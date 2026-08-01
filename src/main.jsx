import React, { useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';
import './styles.css';

const STORAGE_KEY = 'rittenadministratie-fase1-v1';
const STORAGE_BACKUP_KEY = 'rittenadministratie-fase1-v1-safety-copy';
const AUTO_BACKUPS_KEY = 'rittenadministratie-auto-backups-v1';
const MAX_AUTO_BACKUPS = 10;
const DB_NAME = 'rittenadministratie-db';
const DB_STORE = 'rittenadministratie-store';
const DB_DATA_KEY = 'current-data';

const initialVehicle = {
  driverName: 'Theo Verdooren',
  licensePlate: 'GXX71T',
  vehicleName: 'Volvo EX30',
  loanStartDate: '',
  initialMileage: '',
};

const emptyDraft = {
  departurePlace: '',
  plannedArrivalPlace: '',
  fixedDistance: '',
  routeTemplateId: '',
  purpose: '',
  type: 'Zakelijk',
  driverName: 'Theo Verdooren',
  startMileage: '',
};

function pad(value) {
  return String(value).padStart(2, '0');
}

function toDateInput(date = new Date()) {
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function toTime(date = new Date()) {
  return `${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function formatDate(value) {
  if (!value) return '-';
  const [year, month, day] = value.split('-');
  return `${day}-${month}-${year}`;
}

function formatDateTime(value) {
  if (!value) return '-';
  const date = new Date(value);
  return `${formatDate(toDateInput(date))} ${toTime(date)}`;
}

function formatKm(value) {
  const number = Number(value || 0);
  return `${new Intl.NumberFormat('nl-NL').format(number)} km`;
}

function parseKm(value) {
  return Number(String(value).replace(/\./g, '').replace(',', '.')) || 0;
}

function downloadFile(filename, content, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function getPosition() {
  return new Promise((resolve, reject) => {
    if (!navigator.geolocation) {
      reject(new Error('GPS wordt niet ondersteund op dit apparaat.'));
      return;
    }
    navigator.geolocation.getCurrentPosition(resolve, reject, {
      enableHighAccuracy: true,
      maximumAge: 30000,
      timeout: 15000,
    });
  });
}

async function getPlaceFromPosition(position) {
  const { latitude, longitude } = position.coords;
  const fallback = `${latitude.toFixed(5)}, ${longitude.toFixed(5)}`;

  try {
    const params = new URLSearchParams({
      format: 'jsonv2',
      lat: String(latitude),
      lon: String(longitude),
      zoom: '12',
      addressdetails: '1',
    });
    const response = await fetch(`https://nominatim.openstreetmap.org/reverse?${params}`);
    if (!response.ok) return fallback;
    const result = await response.json();
    const address = result.address || {};
    return (
      address.city ||
      address.town ||
      address.village ||
      address.municipality ||
      address.suburb ||
      address.hamlet ||
      result.name ||
      fallback
    );
  } catch {
    return fallback;
  }
}

function getLocationErrorMessage(error) {
  if (error?.code === 1) return 'Locatie is geweigerd. Zet locatietoegang aan voor deze app/browser.';
  if (error?.code === 2) return 'Locatie kon niet worden bepaald. Probeer het buiten of met GPS aan.';
  if (error?.code === 3) return 'Locatie bepalen duurde te lang. Probeer het nog een keer.';
  return error?.message || 'Locatie kon niet worden bepaald.';
}

function sortChronological(rides) {
  return [...rides].sort((a, b) => {
    const dateCompare = `${a.date} ${a.departureTime}`.localeCompare(`${b.date} ${b.departureTime}`);
    return dateCompare || a.number - b.number;
  });
}

function sortByRideNumber(rides) {
  return [...rides].sort((a, b) => Number(a.number) - Number(b.number));
}

function getHighestMileageRide(rides) {
  return rides.reduce((highest, ride) => {
    if (!highest) return ride;
    return Number(ride.endMileage || 0) > Number(highest.endMileage || 0) ? ride : highest;
  }, null);
}

function getDurationText(startedAt) {
  if (!startedAt) return '-';
  const seconds = Math.max(0, Math.floor((Date.now() - new Date(startedAt).getTime()) / 1000));
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const rest = seconds % 60;
  return `${pad(hours)}:${pad(minutes)}:${pad(rest)}`;
}

function getGreeting(date = new Date()) {
  const hour = date.getHours();
  if (hour < 12) return 'Goedemorgen';
  if (hour < 18) return 'Goedemiddag';
  return 'Goedenavond';
}

function getFirstName(name) {
  return (name || 'Theo').trim().split(/\s+/)[0] || 'Theo';
}

function normalizeStoredData(stored) {
  if (!stored || !stored.vehicle || !Array.isArray(stored.rides) || !('activeRide' in stored)) {
    throw new Error('Ongeldige opslag');
  }
  const vehicle = { ...initialVehicle, ...stored.vehicle };
  if (!vehicle.licensePlate) vehicle.licensePlate = initialVehicle.licensePlate;
  if (!vehicle.vehicleName) vehicle.vehicleName = initialVehicle.vehicleName;
  return {
    vehicle,
    rides: Array.isArray(stored.rides) ? stored.rides : [],
    routeTemplates: Array.isArray(stored.routeTemplates) ? stored.routeTemplates : [],
    activeRide: stored.activeRide || null,
  };
}

function dataScore(data) {
  return (data.rides?.length || 0) * 10 + (data.activeRide ? 5 : 0) + (data.routeTemplates?.length || 0);
}

function chooseSafestData(current, candidate) {
  if (!candidate) return current;
  return dataScore(candidate) > dataScore(current) ? candidate : current;
}

function openDatabase() {
  return new Promise((resolve, reject) => {
    if (!('indexedDB' in window)) {
      reject(new Error('IndexedDB niet beschikbaar'));
      return;
    }
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => {
      request.result.createObjectStore(DB_STORE);
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveDataToIndexedDb(data) {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(DB_STORE, 'readwrite');
    transaction.objectStore(DB_STORE).put(data, DB_DATA_KEY);
    transaction.oncomplete = () => {
      db.close();
      resolve();
    };
    transaction.onerror = () => {
      db.close();
      reject(transaction.error);
    };
  });
}

async function loadDataFromIndexedDb() {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(DB_STORE, 'readonly');
    const request = transaction.objectStore(DB_STORE).get(DB_DATA_KEY);
    request.onsuccess = () => {
      db.close();
      try {
        resolve(request.result ? normalizeStoredData(request.result) : null);
      } catch (error) {
        reject(error);
      }
    };
    request.onerror = () => {
      db.close();
      reject(request.error);
    };
  });
}

function saveDataToStorage(data) {
  const serialized = JSON.stringify(data);
  localStorage.setItem(STORAGE_KEY, serialized);
  if ((data.rides?.length || 0) > 0 || data.activeRide) {
    localStorage.setItem(STORAGE_BACKUP_KEY, serialized);
    saveDataToIndexedDb(data).catch(() => {});
  }
}

function getAutoBackups() {
  try {
    const backups = JSON.parse(localStorage.getItem(AUTO_BACKUPS_KEY));
    return Array.isArray(backups) ? backups : [];
  } catch {
    return [];
  }
}

function saveAutoBackups(backups) {
  localStorage.setItem(AUTO_BACKUPS_KEY, JSON.stringify(backups.slice(0, MAX_AUTO_BACKUPS)));
}

function makeBackupName(rideNumber, date = new Date()) {
  return `rittenadministratie-auto-rit-${rideNumber}-${toDateInput(date)}-${toTime(date).replace(':', '')}.json`;
}

function createAutoBackup(nextData, rideNumber, shouldDownload = true) {
  const createdAt = new Date();
  const backup = {
    id: crypto.randomUUID(),
    name: makeBackupName(rideNumber, createdAt),
    rideNumber,
    createdAt: createdAt.toISOString(),
    data: nextData,
  };
  saveAutoBackups([backup, ...getAutoBackups()]);
  if (shouldDownload) {
    downloadFile(backup.name, JSON.stringify(nextData, null, 2), 'application/json');
  }
  return backup;
}

function loadData() {
  const sources = [STORAGE_KEY, STORAGE_BACKUP_KEY];
  for (const key of sources) {
    try {
      const raw = localStorage.getItem(key);
      if (!raw) continue;
      const data = normalizeStoredData(JSON.parse(raw));
      if (key !== STORAGE_KEY) saveDataToStorage(data);
      return data;
    } catch {
      // Probeer de veiligheidskopie voordat we leeg starten.
    }
  }
  return { vehicle: initialVehicle, rides: [], routeTemplates: [], activeRide: null };
}

function validateConnections(rides) {
  const ordered = sortByRideNumber(rides);
  const warnings = [];
  ordered.forEach((ride, index) => {
    const previous = ordered[index - 1];
    if (previous && Number(ride.startMileage) !== Number(previous.endMileage)) {
      warnings.push(
        `Rit ${ride.number} begint met ${formatKm(ride.startMileage)}, maar rit ${previous.number} eindigt met ${formatKm(previous.endMileage)}.`
      );
    }
  });
  return warnings;
}

function App() {
  const [data, setData] = useState(loadData);
  const [draft, setDraft] = useState(emptyDraft);
  const [showStartForm, setShowStartForm] = useState(false);
  const [showManualForm, setShowManualForm] = useState(false);
  const [manualRide, setManualRide] = useState(null);
  const [routeDraft, setRouteDraft] = useState({ name: '', from: '', to: '', distance: '', addReturn: true });
  const [finishDraft, setFinishDraft] = useState({ arrivalPlace: '', endMileage: '' });
  const [editingRide, setEditingRide] = useState(null);
  const [filters, setFilters] = useState({ search: '', from: '', to: '', type: 'Alle' });
  const [message, setMessage] = useState('');
  const [installPrompt, setInstallPrompt] = useState(null);
  const [isOfflineReady, setIsOfflineReady] = useState(false);
  const [locationLoading, setLocationLoading] = useState('');
  const [mobileView, setMobileView] = useState('home');
  const [period, setPeriod] = useState('week');
  const [, setTicker] = useState(0);
  const backupInputRef = useRef(null);

  function updateData(updater) {
    setData((current) => {
      const next = typeof updater === 'function' ? updater(current) : updater;
      saveDataToStorage(next);
      return next;
    });
  }

  useEffect(() => {
    let cancelled = false;
    loadDataFromIndexedDb()
      .then((storedData) => {
        if (cancelled || !storedData) return;
        setData((current) => {
          const safest = chooseSafestData(current, storedData);
          if (safest !== current) {
            saveDataToStorage(safest);
            setMessage('Gegevens hersteld uit veilige telefoonopslag.');
          }
          return safest;
        });
      })
      .catch(() => {});
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    const timer = setInterval(() => setTicker((value) => value + 1), 1000);
    return () => clearInterval(timer);
  }, []);

  useEffect(() => {
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker
        .register('/sw.js')
        .then(() => setIsOfflineReady(true))
        .catch(() => setIsOfflineReady(false));
    }

    const handleInstallPrompt = (event) => {
      event.preventDefault();
      setInstallPrompt(event);
    };
    const handleInstalled = () => {
      setInstallPrompt(null);
      setMessage('App geïnstalleerd. Je kunt hem voortaan vanaf je startscherm openen.');
    };

    window.addEventListener('beforeinstallprompt', handleInstallPrompt);
    window.addEventListener('appinstalled', handleInstalled);
    return () => {
      window.removeEventListener('beforeinstallprompt', handleInstallPrompt);
      window.removeEventListener('appinstalled', handleInstalled);
    };
  }, []);

  useEffect(() => {
    if (!data.activeRide) return;
    setFinishDraft((current) => {
      if (current.arrivalPlace || current.endMileage !== '') return current;
      return {
        arrivalPlace: data.activeRide.plannedArrivalPlace || '',
        endMileage:
          data.activeRide.fixedDistance !== '' && data.activeRide.fixedDistance !== undefined
            ? Number(data.activeRide.startMileage) + Number(data.activeRide.fixedDistance)
            : '',
      };
    });
  }, [data.activeRide]);

  const ridesByNumber = useMemo(() => sortByRideNumber(data.rides), [data.rides]);
  const latestRide = ridesByNumber.at(-1);
  const highestMileageRide = useMemo(() => getHighestMileageRide(data.rides), [data.rides]);
  const latestRides = useMemo(() => [...data.rides].sort((a, b) => b.number - a.number), [data.rides]);
  const nextNumber = Math.max(0, ...data.rides.map((ride) => ride.number)) + 1;
  const lastMileage = highestMileageRide ? highestMileageRide.endMileage : parseKm(data.vehicle.initialMileage);
  const connectionWarnings = useMemo(() => validateConnections(data.rides), [data.rides]);
  const driverFirstName = getFirstName(data.vehicle.driverName);
  const greeting = `${getGreeting()}, ${driverFirstName}`;

  const filteredRides = useMemo(() => {
    const search = filters.search.toLowerCase().trim();
    return [...data.rides]
      .filter((ride) => {
        const text = `${ride.departurePlace} ${ride.arrivalPlace} ${ride.driverName} ${ride.purpose}`.toLowerCase();
        const matchesSearch = !search || text.includes(search);
        const matchesFrom = !filters.from || ride.date >= filters.from;
        const matchesTo = !filters.to || ride.date <= filters.to;
        const matchesType = filters.type === 'Alle' || ride.type === filters.type;
        return matchesSearch && matchesFrom && matchesTo && matchesType;
      })
      .sort((a, b) => b.number - a.number);
  }, [data.rides, filters]);

  const stats = useMemo(() => {
    const today = toDateInput();
    const now = new Date();
    const startOfWeek = new Date(now);
    startOfWeek.setDate(now.getDate() - ((now.getDay() + 6) % 7));
    const weekInput = toDateInput(startOfWeek);
    const monthInput = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-01`;
    const sum = (rides) => rides.reduce((total, ride) => total + Number(ride.kilometers || 0), 0);
    return {
      totalRides: data.rides.length,
      totalKm: sum(data.rides),
      businessKm: sum(data.rides.filter((ride) => ride.type === 'Zakelijk')),
      privateKm: sum(data.rides.filter((ride) => ride.type === 'Prive')),
      currentMileage: data.activeRide ? data.activeRide.startMileage : lastMileage,
      lastRide: latestRide ? `${formatDate(latestRide.date)} ${latestRide.arrivalTime}` : '-',
      todayKm: sum(data.rides.filter((ride) => ride.date === today)),
      weekKm: sum(data.rides.filter((ride) => ride.date >= weekInput)),
      monthKm: sum(data.rides.filter((ride) => ride.date >= monthInput)),
    };
  }, [data.rides, data.activeRide, latestRide, lastMileage]);

  const periodStats = useMemo(() => {
    const today = toDateInput();
    const now = new Date();
    const startOfWeek = new Date(now);
    startOfWeek.setDate(now.getDate() - ((now.getDay() + 6) % 7));
    const weekInput = toDateInput(startOfWeek);
    const monthInput = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-01`;
    const inPeriod = (ride) => {
      if (period === 'today') return ride.date === today;
      if (period === 'month') return ride.date >= monthInput;
      return ride.date >= weekInput;
    };
    const rides = data.rides.filter(inPeriod);
    const business = rides.filter((ride) => ride.type === 'Zakelijk').reduce((total, ride) => total + Number(ride.kilometers || 0), 0);
    const privateKm = rides.filter((ride) => ride.type !== 'Zakelijk').reduce((total, ride) => total + Number(ride.kilometers || 0), 0);
    const total = business + privateKm;
    return {
      business,
      privateKm,
      total,
      businessPercent: total ? Math.round((business / total) * 100) : 0,
      privatePercent: total ? Math.round((privateKm / total) * 100) : 0,
    };
  }, [data.rides, period]);

  function updateVehicle(field, value) {
    updateData((current) => ({ ...current, vehicle: { ...current.vehicle, [field]: value } }));
  }

  async function installApp() {
    if (!installPrompt) {
      setMessage('Gebruik op iPhone/iPad: delen-knop in Safari en daarna "Zet op beginscherm".');
      return;
    }
    installPrompt.prompt();
    await installPrompt.userChoice;
    setInstallPrompt(null);
  }

  function routeLabel(route) {
    return `${route.name || `${route.from} - ${route.to}`} (${formatKm(route.distance)})`;
  }

  async function fillPlaceFromGps(target, setter) {
    if (!window.isSecureContext) {
      setMessage('GPS werkt alleen via de veilige Vercel-link of op localhost.');
      return;
    }
    setLocationLoading(target);
    setMessage('Locatie bepalen...');
    try {
      const position = await getPosition();
      const place = await getPlaceFromPosition(position);
      setter(place);
      const accuracy = position.coords.accuracy ? ` Nauwkeurigheid ongeveer ${Math.round(position.coords.accuracy)} meter.` : '';
      setMessage(`Locatie ingevuld: ${place}.${accuracy}`);
    } catch (error) {
      setMessage(getLocationErrorMessage(error));
    } finally {
      setLocationLoading('');
    }
  }

  function applyRouteToDraft(routeId) {
    const route = data.routeTemplates.find((item) => item.id === routeId);
    if (!route) {
      setDraft((current) => ({ ...current, routeTemplateId: '', plannedArrivalPlace: '', fixedDistance: '' }));
      return;
    }
    setDraft((current) => ({
      ...current,
      routeTemplateId: route.id,
      departurePlace: route.from,
      plannedArrivalPlace: route.to,
      fixedDistance: route.distance,
      purpose: current.purpose || route.name,
    }));
  }

  function applyRouteToManualRide(routeId) {
    const route = data.routeTemplates.find((item) => item.id === routeId);
    if (!route) {
      setManualRide((current) => ({ ...current, routeTemplateId: '', fixedDistance: '' }));
      return;
    }
    setManualRide((current) => {
      const endMileage = current.startMileage !== '' ? parseKm(current.startMileage) + Number(route.distance) : '';
      return {
        ...current,
        routeTemplateId: route.id,
        departurePlace: route.from,
        arrivalPlace: route.to,
        fixedDistance: route.distance,
        endMileage,
        purpose: current.purpose || route.name,
      };
    });
  }

  function saveRouteTemplate(event) {
    event.preventDefault();
    const distance = parseKm(routeDraft.distance);
    if (!routeDraft.from.trim() || !routeDraft.to.trim() || routeDraft.distance === '' || distance <= 0) {
      setMessage('Vul vertrekplaats, aankomstplaats en vaste afstand in voor de veelvoorkomende rit.');
      return;
    }
    const baseRoute = {
      id: crypto.randomUUID(),
      name: routeDraft.name.trim() || `${routeDraft.from.trim()} - ${routeDraft.to.trim()}`,
      from: routeDraft.from.trim(),
      to: routeDraft.to.trim(),
      distance,
    };
    const routesToAdd = [baseRoute];
    if (routeDraft.addReturn) {
      routesToAdd.push({
        id: crypto.randomUUID(),
        name: routeDraft.name.trim() ? `${routeDraft.name.trim()} terug` : `${baseRoute.to} - ${baseRoute.from}`,
        from: baseRoute.to,
        to: baseRoute.from,
        distance,
      });
    }
    updateData((current) => ({ ...current, routeTemplates: [...(current.routeTemplates || []), ...routesToAdd] }));
    setRouteDraft({ name: '', from: '', to: '', distance: '', addReturn: true });
    setMessage('Veelvoorkomende rit opgeslagen.');
  }

  function deleteRouteTemplate(routeId) {
    const route = data.routeTemplates.find((item) => item.id === routeId);
    const ok = window.confirm(`Veelvoorkomende rit "${routeLabel(route)}" verwijderen? Bestaande ritten blijven staan.`);
    if (!ok) return;
    updateData((current) => ({ ...current, routeTemplates: current.routeTemplates.filter((item) => item.id !== routeId) }));
    setMessage('Veelvoorkomende rit verwijderd.');
  }

  function prepareStartRide() {
    if (data.activeRide) {
      setMessage('Er is al een actieve rit. Beëindig die rit eerst.');
      return;
    }
    if (!data.vehicle.initialMileage && !latestRide) {
      setMessage('Vul eerst de eerste kilometerstand in.');
      return;
    }
    const now = new Date();
    setDraft({
      ...emptyDraft,
      driverName: data.vehicle.driverName || 'Theo Verdooren',
      departurePlace: latestRide?.arrivalPlace || '',
      startMileage: lastMileage,
      number: nextNumber,
      date: toDateInput(now),
      departureTime: toTime(now),
      startedAt: now.toISOString(),
    });
    setShowStartForm(true);
    setMessage('');
  }

  function prepareManualRide() {
    const today = toDateInput();
    setManualRide({
      id: crypto.randomUUID(),
      number: nextNumber,
      date: today,
      driverName: data.vehicle.driverName || 'Theo Verdooren',
      departurePlace: latestRide?.arrivalPlace || '',
      arrivalPlace: '',
      departureTime: '',
      arrivalTime: '',
      startMileage: lastMileage || '',
      endMileage: '',
      fixedDistance: '',
      routeTemplateId: '',
      type: 'Zakelijk',
      purpose: '',
      startedAt: '',
      finishedAt: '',
    });
    setShowManualForm(true);
    setMessage('');
  }

  function startRide(event) {
    event.preventDefault();
    if (!draft.driverName || !draft.departurePlace || draft.startMileage === '') {
      setMessage('Vul bestuurder, vertrekplaats en beginstand in voordat de rit start.');
      return;
    }
    if (data.activeRide) {
      setMessage('Er is al een actieve rit. Er kan maar één rit tegelijk actief zijn.');
      return;
    }
    if (data.rides.some((ride) => Number(ride.number) === Number(draft.number))) {
      setMessage(`Ritnummer ${draft.number} bestaat al. Kies een ander ritnummer.`);
      return;
    }
    updateData((current) => ({
      ...current,
      activeRide: {
        id: crypto.randomUUID(),
        number: draft.number,
        date: draft.date,
        departureTime: draft.departureTime,
        startedAt: draft.startedAt,
        driverName: draft.driverName,
        startMileage: parseKm(draft.startMileage),
        departurePlace: draft.departurePlace.trim(),
        plannedArrivalPlace: draft.plannedArrivalPlace.trim(),
        fixedDistance: draft.fixedDistance !== '' ? parseKm(draft.fixedDistance) : '',
        routeTemplateId: draft.routeTemplateId,
        purpose: draft.purpose.trim(),
        type: draft.type,
      },
    }));
    setFinishDraft({
      arrivalPlace: draft.plannedArrivalPlace || '',
      endMileage: draft.fixedDistance !== '' ? parseKm(draft.startMileage) + parseKm(draft.fixedDistance) : '',
    });
    setShowStartForm(false);
    setMessage('Rit gestart. Goede reis.');
  }

  function finishRide(event) {
    event?.preventDefault();
    if (!data.activeRide) {
      setMessage('Er is geen actieve rit om te beëindigen.');
      return;
    }
    const endMileage = parseKm(finishDraft.endMileage);
    if (!finishDraft.arrivalPlace.trim() || finishDraft.endMileage === '') {
      setMessage('Vul aankomstplaats en eindstand in.');
      window.alert('Vul aankomstplaats en eindstand in voordat je de rit beëindigt.');
      return;
    }
    if (endMileage < Number(data.activeRide.startMileage)) {
      setMessage('De eindstand mag niet lager zijn dan de beginstand.');
      window.alert('De eindstand mag niet lager zijn dan de beginstand.');
      return;
    }
    const now = new Date();
    const completed = {
      ...data.activeRide,
      arrivalPlace: finishDraft.arrivalPlace.trim(),
      endMileage,
      arrivalTime: toTime(now),
      finishedAt: now.toISOString(),
      kilometers: endMileage - Number(data.activeRide.startMileage),
    };
    updateData((current) => {
      const next = { ...current, activeRide: null, rides: [...current.rides, completed] };
      createAutoBackup(next, completed.number);
      return next;
    });
    setFinishDraft({ arrivalPlace: '', endMileage: '' });
    setMessage('Rit opgeslagen.');
  }

  function cancelActiveRide() {
    if (!data.activeRide) return;
    const ok = window.confirm(`Actieve rit ${data.activeRide.number} annuleren? Deze rit wordt niet opgeslagen, je eerdere ritten blijven staan.`);
    if (!ok) return;
    updateData((current) => ({ ...current, activeRide: null }));
    setFinishDraft({ arrivalPlace: '', endMileage: '' });
    setMessage('Actieve rit geannuleerd. Je kunt weer een rit starten.');
  }

  function saveManualRide(event) {
    event.preventDefault();
    const ride = {
      ...manualRide,
      number: Number(manualRide.number),
      startMileage: parseKm(manualRide.startMileage),
      endMileage: parseKm(manualRide.endMileage),
    };
    if (
      !ride.number ||
      !ride.date ||
      !ride.driverName ||
      !ride.departurePlace ||
      !ride.arrivalPlace ||
      manualRide.startMileage === '' ||
      manualRide.endMileage === ''
    ) {
      setMessage('Vul alle verplichte velden van de handmatige rit in.');
      return;
    }
    if (ride.endMileage < ride.startMileage) {
      setMessage('De eindstand mag niet lager zijn dan de beginstand.');
      return;
    }
    if (data.rides.some((item) => Number(item.number) === ride.number)) {
      setMessage(`Ritnummer ${ride.number} bestaat al. Kies een ander ritnummer of bewerk de bestaande rit.`);
      return;
    }
    ride.kilometers = ride.endMileage - ride.startMileage;
    updateData((current) => {
      const next = { ...current, rides: [...current.rides, ride] };
      createAutoBackup(next, ride.number);
      return next;
    });
    setShowManualForm(false);
    setManualRide(null);
    setMessage('Handmatige rit opgeslagen.');
  }

  function saveEditedRide(event) {
    event.preventDefault();
    const ride = {
      ...editingRide,
      number: Number(editingRide.number),
      startMileage: parseKm(editingRide.startMileage),
      endMileage: parseKm(editingRide.endMileage),
    };
    if (!ride.number || !ride.driverName || !ride.departurePlace || !ride.arrivalPlace || !ride.date) {
      setMessage('Niet alle verplichte velden zijn ingevuld.');
      return;
    }
    if (ride.endMileage < ride.startMileage) {
      setMessage('De eindstand mag niet lager zijn dan de beginstand.');
      return;
    }
    if (data.rides.some((item) => item.id !== ride.id && Number(item.number) === ride.number)) {
      setMessage(`Ritnummer ${ride.number} bestaat al. Kies een ander ritnummer.`);
      return;
    }
    ride.kilometers = ride.endMileage - ride.startMileage;
    updateData((current) => ({
      ...current,
      rides: current.rides.map((item) => (item.id === ride.id ? ride : item)),
    }));
    setEditingRide(null);
    setMessage('Rit bijgewerkt. Controleer eventuele waarschuwingen over aansluitende kilometerstanden.');
  }

  function deleteRide(ride) {
    const ok = window.confirm(`Rit ${ride.number} verwijderen? Latere kilometerstanden worden niet automatisch aangepast.`);
    if (!ok) return;
    updateData((current) => ({ ...current, rides: current.rides.filter((item) => item.id !== ride.id) }));
    setMessage('Rit verwijderd. Controleer of de kilometerstanden nog aansluiten.');
  }

  function exportPdf() {
    const rides = sortChronological(filteredRides);
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const totals = rides.reduce(
      (acc, ride) => {
        acc.total += ride.kilometers;
        acc[ride.type === 'Zakelijk' ? 'business' : 'private'] += ride.kilometers;
        return acc;
      },
      { total: 0, business: 0, private: 0 }
    );
    doc.setFontSize(16);
    doc.text('Rittenadministratie', 14, 14);
    doc.setFontSize(9);
    doc.text(`Bestuurder: ${data.vehicle.driverName || '-'}`, 14, 22);
    doc.text(`Kenteken: ${data.vehicle.licensePlate || '-'}`, 14, 27);
    doc.text(`Auto: ${data.vehicle.vehicleName || '-'}`, 14, 32);
    doc.text(`Periode: ${filters.from ? formatDate(filters.from) : 'begin'} t/m ${filters.to ? formatDate(filters.to) : 'einde'}`, 90, 22);
    doc.text(`Gemaakt op: ${formatDateTime(new Date().toISOString())}`, 90, 27);
    doc.text(`Totaal: ${formatKm(totals.total)} | Zakelijk: ${formatKm(totals.business)} | Prive: ${formatKm(totals.private)}`, 90, 32);
    autoTable(doc, {
      startY: 39,
      head: [['Rit', 'Datum', 'Bestuurder', 'Van', 'Naar', 'Vertrek', 'Aankomst', 'Beginstand', 'Eindstand', 'Km', 'Type', 'Doel']],
      body: rides.map((ride) => [
        ride.number,
        formatDate(ride.date),
        ride.driverName,
        ride.departurePlace,
        ride.arrivalPlace,
        ride.departureTime,
        ride.arrivalTime,
        formatKm(ride.startMileage),
        formatKm(ride.endMileage),
        ride.kilometers,
        ride.type,
        ride.purpose || '',
      ]),
      styles: { fontSize: 7, cellPadding: 1.5 },
      headStyles: { fillColor: [14, 43, 74] },
    });
    doc.save('rittenadministratie.pdf');
  }

  function exportExcel() {
    const rides = sortChronological(filteredRides);
    const summary = [
      ['Bestuurder', data.vehicle.driverName || ''],
      ['Kenteken', data.vehicle.licensePlate || ''],
      ['Voertuig', data.vehicle.vehicleName || ''],
      ['Periode', `${filters.from ? formatDate(filters.from) : 'begin'} t/m ${filters.to ? formatDate(filters.to) : 'einde'}`],
      ['Aantal ritten', rides.length],
      ['Totaal aantal kilometers', rides.reduce((sum, ride) => sum + ride.kilometers, 0)],
      ['Zakelijke kilometers', rides.filter((ride) => ride.type === 'Zakelijk').reduce((sum, ride) => sum + ride.kilometers, 0)],
      ['Privékilometers', rides.filter((ride) => ride.type !== 'Zakelijk').reduce((sum, ride) => sum + ride.kilometers, 0)],
    ];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(summary), 'Samenvatting');
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(
        rides.map((ride) => ({
          Rit: ride.number,
          Datum: formatDate(ride.date),
          Bestuurder: ride.driverName,
          Van: ride.departurePlace,
          Naar: ride.arrivalPlace,
          Vertrek: ride.departureTime,
          Aankomst: ride.arrivalTime,
          Beginstand: ride.startMileage,
          Eindstand: ride.endMileage,
          Kilometers: ride.kilometers,
          Type: ride.type,
          Doel: ride.purpose || '',
        }))
      ),
      'Ritten'
    );
    XLSX.writeFile(workbook, 'rittenadministratie.xlsx');
  }

  function saveBackup() {
    downloadFile(`rittenadministratie-backup-${toDateInput()}.json`, JSON.stringify(data, null, 2), 'application/json');
  }

  function downloadLatestAutoBackup() {
    const [latest] = getAutoBackups();
    if (!latest) {
      setMessage('Er is nog geen automatische back-up.');
      return;
    }
    downloadFile(latest.name, JSON.stringify(latest.data, null, 2), 'application/json');
    setMessage(`Automatische back-up gedownload: ${latest.name}`);
  }

  function restoreLatestAutoBackup() {
    const [latest] = getAutoBackups();
    if (!latest) {
      setMessage('Er is nog geen automatische back-up om terug te zetten.');
      return;
    }
    const ok = window.confirm(`Laatste automatische back-up terugzetten?\n\n${latest.name}\n\nJe huidige gegevens worden vervangen.`);
    if (!ok) return;
    updateData(normalizeStoredData(latest.data));
    setMessage(`Automatische back-up teruggezet: ${latest.name}`);
  }

  function restoreBackup(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        if (!parsed.vehicle || !Array.isArray(parsed.rides) || !('activeRide' in parsed)) {
          throw new Error('Ongeldig bestand');
        }
        const ok = window.confirm('Back-up terugzetten? De huidige gegevens worden vervangen.');
        if (!ok) return;
        updateData({
          vehicle: { ...initialVehicle, ...parsed.vehicle },
          rides: parsed.rides,
          routeTemplates: Array.isArray(parsed.routeTemplates) ? parsed.routeTemplates : [],
          activeRide: parsed.activeRide,
        });
        setMessage('Back-up teruggezet.');
      } catch {
        setMessage('Dit back-upbestand kon niet worden gelezen of is ongeldig.');
      } finally {
        backupInputRef.current.value = '';
      }
    };
    reader.readAsText(file);
  }

  return (
    <main data-mobile-view={mobileView}>
      <header className="topbar app-header">
        <AppLogo />
        <div className="top-actions">
          {installPrompt && <button className="icon-button install-button" onClick={installApp} title="App installeren">Installeren</button>}
          <button className="icon-button" onClick={() => setMobileView('export')} title="Export en back-up">Export</button>
          <button className="icon-button" onClick={() => setMobileView('more')} title="Instellingen">Instellingen</button>
          <input ref={backupInputRef} type="file" accept="application/json" hidden onChange={(event) => restoreBackup(event.target.files[0])} />
        </div>
      </header>

      <section className="hero-panel" data-mobile-section="home">
        <h1>{greeting}</h1>
        <VehicleCard vehicle={data.vehicle} currentMileage={stats.currentMileage} />
      </section>

      {message && <div className="notice">{message}</div>}
      {connectionWarnings.length > 0 && (
        <section className="warning">
          <strong>Kilometerstanden controleren</strong>
          {connectionWarnings.map((warning) => (
            <p key={warning}>{warning}</p>
          ))}
        </section>
      )}

      <Dashboard stats={stats} period={period} setPeriod={setPeriod} periodStats={periodStats} />

      <section className="layout app-section" data-mobile-section="home">
        <VehicleSettings vehicle={data.vehicle} updateVehicle={updateVehicle} />
        <RideControl
          activeRide={data.activeRide}
          draft={draft}
          setDraft={setDraft}
          showStartForm={showStartForm}
          setShowStartForm={setShowStartForm}
          prepareStartRide={prepareStartRide}
          startRide={startRide}
          finishDraft={finishDraft}
          setFinishDraft={setFinishDraft}
          finishRide={finishRide}
          cancelActiveRide={cancelActiveRide}
          routeTemplates={data.routeTemplates}
          routeLabel={routeLabel}
          applyRouteToDraft={applyRouteToDraft}
          fillPlaceFromGps={fillPlaceFromGps}
          locationLoading={locationLoading}
        />
      </section>

      <section className="manual-shortcut app-section" data-mobile-section="home">
        <button
          type="button"
          onClick={() => {
            prepareManualRide();
            setMobileView('more');
          }}
        >
          Handmatig invullen
        </button>
      </section>

      <LatestRides rides={latestRides} onEdit={setEditingRide} onViewAll={() => setMobileView('rides')} />

      <RouteTemplates
        routes={data.routeTemplates}
        routeDraft={routeDraft}
        setRouteDraft={setRouteDraft}
        saveRouteTemplate={saveRouteTemplate}
        deleteRouteTemplate={deleteRouteTemplate}
        routeLabel={routeLabel}
      />

      <ManualRideEntry
        showManualForm={showManualForm}
        manualRide={manualRide}
        setManualRide={setManualRide}
        prepareManualRide={prepareManualRide}
        saveManualRide={saveManualRide}
        routeTemplates={data.routeTemplates}
        routeLabel={routeLabel}
        applyRouteToManualRide={applyRouteToManualRide}
        cancelManualRide={() => {
          setShowManualForm(false);
          setManualRide(null);
        }}
      />

      <ExportPanel
        saveBackup={saveBackup}
        restoreBackup={() => backupInputRef.current.click()}
        exportPdf={exportPdf}
        exportExcel={exportExcel}
        downloadLatestAutoBackup={downloadLatestAutoBackup}
        restoreLatestAutoBackup={restoreLatestAutoBackup}
      />

      <section className="panel overview-panel app-section" data-mobile-section="rides">
        <div className="section-heading">
          <div>
            <h2>Rittenoverzicht</h2>
            <p>Nieuwste ritten staan bovenaan. Export gebruikt chronologische volgorde.</p>
          </div>
          <div className="actions">
            <button onClick={exportPdf}>Exporteren naar PDF</button>
            <button onClick={exportExcel}>Exporteren naar Excel</button>
          </div>
        </div>
        <Filters filters={filters} setFilters={setFilters} />
        <RideTable rides={filteredRides} onEdit={setEditingRide} onDelete={deleteRide} />
      </section>

      {editingRide && (
        <EditRideModal ride={editingRide} setRide={setEditingRide} onSave={saveEditedRide} onClose={() => setEditingRide(null)} />
      )}

      <MobileNav activeView={mobileView} setActiveView={setMobileView} />

      <footer>
        Alles wordt alleen in deze browser opgeslagen. Wanneer browsergegevens worden gewist, kunnen ritten verdwijnen. Maak daarom regelmatig een back-up of export.
      </footer>
    </main>
  );
}

function AppLogo() {
  return (
    <div className="app-logo" aria-label="THEO Rittenadministratie">
      <div className="logo-mark">
        <span className="logo-road" />
        <span className="logo-pin" />
      </div>
      <div>
        <strong>THEO</strong>
        <span>Rittenadministratie</span>
      </div>
    </div>
  );
}

function VehicleCard({ vehicle, currentMileage }) {
  return (
    <article className="vehicle-card">
      <div>
        <h2>{vehicle.vehicleName || 'Voertuig'}</h2>
        <span className="license-chip">{vehicle.licensePlate || '-'}</span>
        <strong>{formatKm(currentMileage)}</strong>
        <p>Laatste stand</p>
      </div>
      <div className="road-orb" aria-hidden="true">
        <span />
      </div>
    </article>
  );
}

function LatestRides({ rides, onEdit, onViewAll }) {
  return (
    <section className="panel latest-panel app-section" data-mobile-section="home">
      <div className="section-heading compact-heading">
        <h2>Laatste ritten</h2>
        <button className="text-button" type="button" onClick={onViewAll}>Bekijk alles</button>
      </div>
      <div className="latest-list">
        {rides.length === 0 && <p className="muted">Nog geen ritten opgeslagen.</p>}
        {rides.map((ride) => (
          <button className="latest-item" type="button" key={ride.id} onClick={() => onEdit({ ...ride })}>
            <span className={`ride-type-icon ${ride.type === 'Prive' ? 'private' : 'business'}`} />
            <span>
              <strong>{ride.departurePlace} → {ride.arrivalPlace}</strong>
              <small>{formatDate(ride.date)}</small>
            </span>
            <span className="latest-km">
              <strong>{formatKm(ride.kilometers)}</strong>
              <small className={ride.type === 'Prive' ? 'private-text' : 'business-text'}>{ride.type === 'Prive' ? 'Privé' : 'Zakelijk'}</small>
            </span>
            <span className="chevron">›</span>
          </button>
        ))}
      </div>
    </section>
  );
}

function ExportPanel({ saveBackup, restoreBackup, exportPdf, exportExcel, downloadLatestAutoBackup, restoreLatestAutoBackup }) {
  return (
    <section className="panel export-panel app-section" data-mobile-section="export">
      <div className="section-heading">
        <div>
          <h2>Export en back-up</h2>
          <p>Bewaar of herstel je rittenadministratie.</p>
        </div>
      </div>
      <div className="export-grid">
        <button onClick={exportPdf}>PDF exporteren</button>
        <button onClick={exportExcel}>Excel exporteren</button>
        <button onClick={saveBackup}>Back-up opslaan</button>
        <button onClick={restoreBackup}>Back-up terugzetten</button>
        <button onClick={downloadLatestAutoBackup}>Laatste auto-back-up downloaden</button>
        <button onClick={restoreLatestAutoBackup}>Laatste auto-back-up terugzetten</button>
      </div>
    </section>
  );
}

function MobileNav({ activeView, setActiveView }) {
  const items = [
    ['home', 'Home'],
    ['rides', 'Ritten'],
    ['export', 'Export'],
    ['more', 'Meer'],
  ];
  return (
    <nav className="mobile-nav" aria-label="Mobiele navigatie">
      {items.map(([value, label]) => (
        <button
          key={value}
          type="button"
          className={activeView === value ? 'active' : ''}
          onClick={() => setActiveView(value)}
        >
          <span className={`nav-icon ${value}`} />
          {label}
        </button>
      ))}
    </nav>
  );
}

function VehicleSettings({ vehicle, updateVehicle }) {
  return (
    <section className="panel vehicle-panel" data-mobile-section="more">
      <h2>Voertuiggegevens</h2>
      <div className="form-grid">
        <label>
          Naam bestuurder
          <input value={vehicle.driverName} onChange={(event) => updateVehicle('driverName', event.target.value)} />
        </label>
        <label>
          Kenteken
          <input value={vehicle.licensePlate} onChange={(event) => updateVehicle('licensePlate', event.target.value.toUpperCase())} />
        </label>
        <label>
          Merk en type auto
          <input value={vehicle.vehicleName} onChange={(event) => updateVehicle('vehicleName', event.target.value)} />
        </label>
        <label>
          Start leenperiode
          <input type="date" value={vehicle.loanStartDate} onChange={(event) => updateVehicle('loanStartDate', event.target.value)} />
        </label>
        <label>
          Eerste kilometerstand
          <input inputMode="numeric" value={vehicle.initialMileage} onChange={(event) => updateVehicle('initialMileage', event.target.value)} />
        </label>
      </div>
    </section>
  );
}

function RideControl(props) {
  const {
    activeRide,
    draft,
    setDraft,
    showStartForm,
    setShowStartForm,
    prepareStartRide,
    startRide,
    finishDraft,
    setFinishDraft,
    finishRide,
    cancelActiveRide,
    routeTemplates,
    routeLabel,
    applyRouteToDraft,
    fillPlaceFromGps,
    locationLoading,
  } = props;

  return (
    <section className="panel ride-panel" data-mobile-section="home">
      <h2>Rit</h2>
      {!activeRide && !showStartForm && (
        <button className="primary-start" onClick={prepareStartRide}>
          Rit starten
        </button>
      )}
      {!activeRide && showStartForm && (
        <form onSubmit={startRide}>
          <div className="form-grid">
            <label className="wide-field">
              Veelvoorkomende rit
              <select value={draft.routeTemplateId} onChange={(event) => applyRouteToDraft(event.target.value)}>
                <option value="">Geen vaste route</option>
                {routeTemplates.map((route) => (
                  <option key={route.id} value={route.id}>{routeLabel(route)}</option>
                ))}
              </select>
            </label>
            <label>
              Ritnummer
              <input value={draft.number} onChange={(event) => setDraft({ ...draft, number: Number(event.target.value) })} />
            </label>
            <label>
              Datum
              <input type="date" value={draft.date} onChange={(event) => setDraft({ ...draft, date: event.target.value })} />
            </label>
            <label>
              Vertrektijd
              <input type="time" value={draft.departureTime} onChange={(event) => setDraft({ ...draft, departureTime: event.target.value })} />
            </label>
            <label>
              Bestuurder
              <input value={draft.driverName} onChange={(event) => setDraft({ ...draft, driverName: event.target.value })} />
            </label>
            <label>
              Beginstand
              <input inputMode="numeric" value={draft.startMileage} onChange={(event) => setDraft({ ...draft, startMileage: event.target.value })} />
            </label>
            <label>
              Vertrekplaats
              <div className="input-with-button">
                <input value={draft.departurePlace} onChange={(event) => setDraft({ ...draft, departurePlace: event.target.value })} />
                <button
                  type="button"
                  onClick={() => fillPlaceFromGps('departure', (place) => setDraft((current) => ({ ...current, departurePlace: place })))}
                  disabled={locationLoading === 'departure'}
                >
                  {locationLoading === 'departure' ? 'Zoeken...' : 'Gebruik GPS'}
                </button>
              </div>
            </label>
            <label>
              Geplande aankomstplaats
              <input value={draft.plannedArrivalPlace} onChange={(event) => setDraft({ ...draft, plannedArrivalPlace: event.target.value })} />
            </label>
            <label>
              Vaste afstand
              <input inputMode="numeric" value={draft.fixedDistance} onChange={(event) => setDraft({ ...draft, fixedDistance: event.target.value })} />
            </label>
            <label>
              Type
              <select value={draft.type} onChange={(event) => setDraft({ ...draft, type: event.target.value })}>
                <option>Zakelijk</option>
                <option value="Prive">Privé</option>
              </select>
            </label>
            <label>
              Doel of reden
              <input value={draft.purpose} onChange={(event) => setDraft({ ...draft, purpose: event.target.value })} />
            </label>
          </div>
          <div className="actions">
            <button className="primary-start" type="submit">Rit definitief starten</button>
            <button type="button" onClick={() => setShowStartForm(false)}>Annuleren</button>
          </div>
        </form>
      )}
      {activeRide && (
        <div className="active-ride">
          <dl>
            <div><dt>Ritnummer</dt><dd>{activeRide.number}</dd></div>
            <div><dt>Datum</dt><dd>{formatDate(activeRide.date)}</dd></div>
            <div><dt>Vertrekplaats</dt><dd>{activeRide.departurePlace}</dd></div>
            <div><dt>Vertrektijd</dt><dd>{activeRide.departureTime}</dd></div>
            <div><dt>Beginstand</dt><dd>{formatKm(activeRide.startMileage)}</dd></div>
            <div><dt>Bestuurder</dt><dd>{activeRide.driverName}</dd></div>
            {activeRide.plannedArrivalPlace && <div><dt>Gepland naar</dt><dd>{activeRide.plannedArrivalPlace}</dd></div>}
            {activeRide.fixedDistance !== '' && <div><dt>Vaste afstand</dt><dd>{formatKm(activeRide.fixedDistance)}</dd></div>}
            <div><dt>Actief</dt><dd>{getDurationText(activeRide.startedAt)}</dd></div>
          </dl>
          <div className="finish-form">
            <label>
              Aankomstplaats
              <div className="input-with-button">
                <input value={finishDraft.arrivalPlace} onChange={(event) => setFinishDraft({ ...finishDraft, arrivalPlace: event.target.value })} />
                <button
                  type="button"
                  onClick={() => fillPlaceFromGps('arrival', (place) => setFinishDraft((current) => ({ ...current, arrivalPlace: place })))}
                  disabled={locationLoading === 'arrival'}
                >
                  {locationLoading === 'arrival' ? 'Zoeken...' : 'Gebruik GPS'}
                </button>
              </div>
            </label>
            <label>
              Actuele eindstand
              <input inputMode="numeric" value={finishDraft.endMileage} onChange={(event) => setFinishDraft({ ...finishDraft, endMileage: event.target.value })} />
            </label>
            <button className="primary-stop" type="button" onClick={finishRide}>Rit beëindigen</button>
            <button className="danger" type="button" onClick={cancelActiveRide}>Actieve rit annuleren</button>
          </div>
        </div>
      )}
    </section>
  );
}

function RouteTemplates({ routes, routeDraft, setRouteDraft, saveRouteTemplate, deleteRouteTemplate, routeLabel }) {
  return (
    <section className="panel route-panel" data-mobile-section="more">
      <div className="section-heading">
        <div>
          <h2>Veelvoorkomende ritten</h2>
          <p>Sla vaste ritten op, zoals Made - Oostvoorne, zodat afstand en plaatsen meteen worden ingevuld.</p>
        </div>
      </div>
      <form className="route-form" onSubmit={saveRouteTemplate}>
        <label>
          Naam
          <input value={routeDraft.name} placeholder="Bijvoorbeeld: Made - Oostvoorne" onChange={(event) => setRouteDraft({ ...routeDraft, name: event.target.value })} />
        </label>
        <label>
          Van
          <input value={routeDraft.from} onChange={(event) => setRouteDraft({ ...routeDraft, from: event.target.value })} />
        </label>
        <label>
          Naar
          <input value={routeDraft.to} onChange={(event) => setRouteDraft({ ...routeDraft, to: event.target.value })} />
        </label>
        <label>
          Afstand
          <input inputMode="numeric" value={routeDraft.distance} onChange={(event) => setRouteDraft({ ...routeDraft, distance: event.target.value })} />
        </label>
        <label className="checkbox-label">
          <input type="checkbox" checked={routeDraft.addReturn} onChange={(event) => setRouteDraft({ ...routeDraft, addReturn: event.target.checked })} />
          Terugweg ook toevoegen
        </label>
        <button type="submit">Veelvoorkomende rit opslaan</button>
      </form>
      <div className="route-list">
        {routes.length === 0 && <p className="muted">Nog geen veelvoorkomende ritten opgeslagen.</p>}
        {routes.map((route) => (
          <div className="route-item" key={route.id}>
            <strong>{routeLabel(route)}</strong>
            <span>{route.from} naar {route.to}</span>
            <button className="danger" onClick={() => deleteRouteTemplate(route.id)}>Verwijderen</button>
          </div>
        ))}
      </div>
    </section>
  );
}

function ManualRideEntry({
  showManualForm,
  manualRide,
  setManualRide,
  prepareManualRide,
  saveManualRide,
  routeTemplates,
  routeLabel,
  applyRouteToManualRide,
  cancelManualRide,
}) {
  return (
    <section className="panel manual-panel" data-mobile-section="more">
      <div className="section-heading">
        <div>
          <h2>Papieren rit handmatig invoeren</h2>
          <p>Gebruik dit voor ritten die je eerder al op papier hebt bijgehouden.</p>
        </div>
        {!showManualForm && <button onClick={prepareManualRide}>Handmatige rit toevoegen</button>}
      </div>
      {showManualForm && manualRide && (
        <form onSubmit={saveManualRide}>
          <div className="form-grid manual-grid">
            <label className="wide-field">
              Veelvoorkomende rit
              <select value={manualRide.routeTemplateId} onChange={(event) => applyRouteToManualRide(event.target.value)}>
                <option value="">Geen vaste route</option>
                {routeTemplates.map((route) => (
                  <option key={route.id} value={route.id}>{routeLabel(route)}</option>
                ))}
              </select>
            </label>
            <label>
              Ritnummer
              <input value={manualRide.number} onChange={(event) => setManualRide({ ...manualRide, number: event.target.value })} />
            </label>
            <label>
              Datum
              <input type="date" value={manualRide.date} onChange={(event) => setManualRide({ ...manualRide, date: event.target.value })} />
            </label>
            <label>
              Bestuurder
              <input value={manualRide.driverName} onChange={(event) => setManualRide({ ...manualRide, driverName: event.target.value })} />
            </label>
            <label>
              Van
              <input value={manualRide.departurePlace} onChange={(event) => setManualRide({ ...manualRide, departurePlace: event.target.value })} />
            </label>
            <label>
              Naar
              <input value={manualRide.arrivalPlace} onChange={(event) => setManualRide({ ...manualRide, arrivalPlace: event.target.value })} />
            </label>
            <label>
              Beginstand
              <input
                inputMode="numeric"
                value={manualRide.startMileage}
                onChange={(event) => {
                  const startMileage = event.target.value;
                  const endMileage = manualRide.fixedDistance !== '' ? parseKm(startMileage) + parseKm(manualRide.fixedDistance) : manualRide.endMileage;
                  setManualRide({ ...manualRide, startMileage, endMileage });
                }}
              />
            </label>
            <label>
              Eindstand
              <input inputMode="numeric" value={manualRide.endMileage} onChange={(event) => setManualRide({ ...manualRide, endMileage: event.target.value })} />
            </label>
            <label>
              Vaste afstand
              <input inputMode="numeric" value={manualRide.fixedDistance} onChange={(event) => setManualRide({ ...manualRide, fixedDistance: event.target.value })} />
            </label>
            <label>
              Type
              <select value={manualRide.type} onChange={(event) => setManualRide({ ...manualRide, type: event.target.value })}>
                <option>Zakelijk</option>
                <option value="Prive">Privé</option>
              </select>
            </label>
            <label className="wide-field">
              Doel of reden
              <input value={manualRide.purpose} onChange={(event) => setManualRide({ ...manualRide, purpose: event.target.value })} />
            </label>
          </div>
          <div className="actions form-actions">
            <button className="primary-start" type="submit">Handmatige rit opslaan</button>
            <button type="button" onClick={cancelManualRide}>Annuleren</button>
          </div>
        </form>
      )}
    </section>
  );
}

function Dashboard({ stats, period, setPeriod, periodStats }) {
  const cards = [
    ['Totaal ritten', stats.totalRides],
    ['Totaal gereden', formatKm(stats.totalKm)],
    ['Vanaf start rijden', formatKm(stats.totalKm)],
    ['Zakelijk', formatKm(stats.businessKm)],
    ['Privé', formatKm(stats.privateKm)],
  ];
  return (
    <section className="dashboard app-section" data-mobile-section="home">
      <div className="section-heading compact-heading dashboard-heading">
        <h2>Overzicht</h2>
      </div>
      <div className="stat-grid">
        {cards.map(([label, value]) => {
          const isPrivate = label.startsWith('Priv');
          return (
          <article className={`stat ${isPrivate ? 'private' : 'business'}`} key={label}>
            <span className={`stat-icon ${isPrivate ? 'private' : 'business'}`} />
            <span>{label}</span>
            <strong>{value}</strong>
          </article>
          );
        })}
      </div>
      <div className="period-card">
        <div className="segment-control">
          <button className={period === 'today' ? 'active' : ''} type="button" onClick={() => setPeriod('today')}>Vandaag</button>
          <button className={period === 'week' ? 'active' : ''} type="button" onClick={() => setPeriod('week')}>Week</button>
          <button className={period === 'month' ? 'active' : ''} type="button" onClick={() => setPeriod('month')}>Maand</button>
        </div>
        <div className="split-row">
          <span>Zakelijk<br /><strong>{formatKm(periodStats.business)}</strong></span>
          <strong>{periodStats.total ? `${periodStats.businessPercent}%` : '0%'}</strong>
          <span>Privé<br /><strong>{formatKm(periodStats.privateKm)}</strong></span>
        </div>
        <div className="split-bar" aria-label="Kilometerverdeling">
          <span style={{ width: `${periodStats.businessPercent}%` }} />
          <i style={{ width: `${periodStats.privatePercent}%` }} />
        </div>
      </div>
    </section>
  );
}

function Filters({ filters, setFilters }) {
  return (
    <div className="filters">
      <label>
        Zoeken
        <input placeholder="Plaats, bestuurder of doel" value={filters.search} onChange={(event) => setFilters({ ...filters, search: event.target.value })} />
      </label>
      <label>
        Vanaf
        <input type="date" value={filters.from} onChange={(event) => setFilters({ ...filters, from: event.target.value })} />
      </label>
      <label>
        Tot en met
        <input type="date" value={filters.to} onChange={(event) => setFilters({ ...filters, to: event.target.value })} />
      </label>
      <label>
        Type
        <select value={filters.type} onChange={(event) => setFilters({ ...filters, type: event.target.value })}>
          <option>Alle</option>
          <option>Zakelijk</option>
          <option value="Prive">Privé</option>
        </select>
      </label>
    </div>
  );
}

function RideTable({ rides, onEdit, onDelete }) {
  const [openRideId, setOpenRideId] = useState('');

  return (
    <>
    <div className="table-wrap desktop-table">
      <table>
        <thead>
          <tr>
            <th>Rit</th>
            <th>Datum</th>
            <th>Bestuurder</th>
            <th>Van</th>
            <th>Naar</th>
            <th>Vertrek</th>
            <th>Aankomst</th>
            <th>Beginstand</th>
            <th>Eindstand</th>
            <th>Kilometers</th>
            <th>Type</th>
            <th>Doel</th>
            <th>Acties</th>
          </tr>
        </thead>
        <tbody>
          {rides.length === 0 && (
            <tr>
              <td colSpan="13" className="empty">Nog geen ritten gevonden.</td>
            </tr>
          )}
          {rides.map((ride) => (
            <tr key={ride.id}>
              <td data-label="Rit">{ride.number}</td>
              <td data-label="Datum">{formatDate(ride.date)}</td>
              <td data-label="Bestuurder">{ride.driverName}</td>
              <td data-label="Van">{ride.departurePlace}</td>
              <td data-label="Naar">{ride.arrivalPlace}</td>
              <td data-label="Vertrek">{ride.departureTime}</td>
              <td data-label="Aankomst">{ride.arrivalTime}</td>
              <td data-label="Beginstand" className="number">{formatKm(ride.startMileage)}</td>
              <td data-label="Eindstand" className="number">{formatKm(ride.endMileage)}</td>
              <td data-label="Kilometers" className="number">{formatKm(ride.kilometers)}</td>
              <td data-label="Type">{ride.type === 'Prive' ? 'Privé' : ride.type}</td>
              <td data-label="Doel">{ride.purpose}</td>
              <td data-label="Acties" className="row-actions">
                <button onClick={() => onEdit({ ...ride })}>Bekijken/bewerken</button>
                <button className="danger" onClick={() => onDelete(ride)}>Verwijderen</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
    <div className="mobile-ride-list">
      {rides.length === 0 && <p className="empty">Nog geen ritten gevonden.</p>}
      {rides.map((ride) => {
        const isOpen = openRideId === ride.id;
        return (
          <article className="mobile-ride-card" key={ride.id}>
            <button
              className="mobile-ride-summary"
              type="button"
              aria-expanded={isOpen}
              onClick={() => setOpenRideId(isOpen ? '' : ride.id)}
            >
              <span>Rit {ride.number}</span>
              <strong>{formatDate(ride.date)}</strong>
            </button>
            {isOpen && (
              <div className="mobile-ride-details">
                <div><span>Bestuurder</span><strong>{ride.driverName}</strong></div>
                <div><span>Van</span><strong>{ride.departurePlace}</strong></div>
                <div><span>Naar</span><strong>{ride.arrivalPlace}</strong></div>
                <div><span>Vertrek</span><strong>{ride.departureTime || '-'}</strong></div>
                <div><span>Aankomst</span><strong>{ride.arrivalTime || '-'}</strong></div>
                <div><span>Beginstand</span><strong>{formatKm(ride.startMileage)}</strong></div>
                <div><span>Eindstand</span><strong>{formatKm(ride.endMileage)}</strong></div>
                <div><span>Kilometers</span><strong>{formatKm(ride.kilometers)}</strong></div>
                <div><span>Type</span><strong>{ride.type === 'Prive' ? 'Privé' : ride.type}</strong></div>
                <div><span>Doel</span><strong>{ride.purpose || '-'}</strong></div>
                <div className="mobile-ride-actions">
                  <button onClick={() => onEdit({ ...ride })}>Bekijken/bewerken</button>
                  <button className="danger" onClick={() => onDelete(ride)}>Verwijderen</button>
                </div>
              </div>
            )}
          </article>
        );
      })}
    </div>
    </>
  );
}

function EditRideModal({ ride, setRide, onSave, onClose }) {
  return (
    <div className="modal-backdrop">
      <form className="modal" onSubmit={onSave}>
        <div className="section-heading">
          <h2>Rit {ride.number} bekijken of bewerken</h2>
          <button type="button" onClick={onClose}>Sluiten</button>
        </div>
        <div className="form-grid">
          <label>Ritnummer<input value={ride.number} onChange={(event) => setRide({ ...ride, number: event.target.value })} /></label>
          <label>Datum<input type="date" value={ride.date} onChange={(event) => setRide({ ...ride, date: event.target.value })} /></label>
          <label>Bestuurder<input value={ride.driverName} onChange={(event) => setRide({ ...ride, driverName: event.target.value })} /></label>
          <label>Van<input value={ride.departurePlace} onChange={(event) => setRide({ ...ride, departurePlace: event.target.value })} /></label>
          <label>Naar<input value={ride.arrivalPlace} onChange={(event) => setRide({ ...ride, arrivalPlace: event.target.value })} /></label>
          <label>Vertrek<input type="time" value={ride.departureTime} onChange={(event) => setRide({ ...ride, departureTime: event.target.value })} /></label>
          <label>Aankomst<input type="time" value={ride.arrivalTime} onChange={(event) => setRide({ ...ride, arrivalTime: event.target.value })} /></label>
          <label>Beginstand<input inputMode="numeric" value={ride.startMileage} onChange={(event) => setRide({ ...ride, startMileage: event.target.value })} /></label>
          <label>Eindstand<input inputMode="numeric" value={ride.endMileage} onChange={(event) => setRide({ ...ride, endMileage: event.target.value })} /></label>
          <label>Type<select value={ride.type} onChange={(event) => setRide({ ...ride, type: event.target.value })}><option>Zakelijk</option><option value="Prive">Privé</option></select></label>
          <label>Doel<input value={ride.purpose || ''} onChange={(event) => setRide({ ...ride, purpose: event.target.value })} /></label>
        </div>
        <button className="primary-start" type="submit">Wijzigingen opslaan</button>
      </form>
    </div>
  );
}

createRoot(document.getElementById('root')).render(<App />);
