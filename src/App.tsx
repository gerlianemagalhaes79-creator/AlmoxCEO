import { DevolutionTab } from './components/tabs/DevolutionTab';
import * as React from 'react';
import { useState, useEffect, useMemo, useCallback, useRef } from 'react';
import { 
  Package, 
  ArrowUpRight, 
  ArrowDownLeft, 
  AlertTriangle, 
  Plus, 
  History, 
  LayoutDashboard,
  Calendar,
  Search,
  Settings,
  ChevronRight,
  Menu,
  X,
  Check,
  Edit2,
  BarChart3,
  TrendingUp,
  Upload,
  TrendingDown,
  DollarSign,
  Filter,
  Download,
  FileText,
  LogIn,
  LogOut,
  Trash2,
  Save,
  RotateCcw,
  CheckCircle,
  Clock,
  Bell,
  Users,
  Info,
  Printer,
  Copy,
  BookOpen,
  Activity,
  PieChart as PieChartIcon,
  Image as ImageIcon,
  Tag
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import html2canvas from 'html2canvas';
import { 
  collection, 
  onSnapshot, 
  addDoc, 
  updateDoc, 
  setDoc,
  getDoc,
  doc, 
  query, 
  orderBy, 
  serverTimestamp,
  runTransaction,
  where,
  Timestamp,
  getDocs,
  writeBatch,
  deleteDoc,
  deleteField,
  DocumentReference
} from 'firebase/firestore';
import { 
  signInWithPopup, 
  signInWithRedirect,
  GoogleAuthProvider, 
  onAuthStateChanged, 
  signOut,
  User,
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  updateProfile,
  getAuth
} from 'firebase/auth';
import { initializeApp } from 'firebase/app';
import { db, auth } from './firebase';
import firebaseConfig from '../firebase-applet-config.json';
import { Item, Transaction, UserProfile, MaterialRequest, RequestItem, Notification } from './types';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer, 
  PieChart, 
  Pie, 
  Cell,
  LineChart,
  Line,
  AreaChart,
  Area,
  Legend
} from 'recharts';
import { format, subDays, isWithinInterval, startOfDay, endOfDay, parseISO, differenceInDays } from 'date-fns';
import { ptBR } from 'date-fns/locale';

interface ItemGroup {
  name: string;
  total_quantity: number;
  min_quantity: number;
  category: string | null;
  supplier: string | null;
  unit_measure?: string | null;
  batches: Item[];
  weeklyExitRate: number;
  durationWeeks: number | 'infinite';
}

const normalizeString = (str: string | null | undefined) => 
  (str || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();

const getSafeDocId = (name: string | null | undefined) => {
  const normalized = normalizeString(name);
  return normalized.replace(/[^a-z0-9]/gi, '_');
};

const SECTORS = [
  'CPSMS', 'CME', 'ClÃ­nica Geral', 'HigienizaÃ§Ã£o', 'DireÃ§Ã£o', 
  'RecepÃ§Ã£o', 'SAME', 'Copa', 'Administrativo', 'TI', 'RegulaÃ§Ã£o'
];

const SECTOR_COLORS: Record<string, string> = {
  'CPSMS': '#0284c7',
  'CME': '#7c3aed',
  'ClÃ­nica Geral': '#059669',
  'HigienizaÃ§Ã£o': '#6366f1',
  'DireÃ§Ã£o': '#ef4444',
  'RecepÃ§Ã£o': '#14b8a6',
  'SAME': '#7c2d12',
  'Copa': '#84cc16',
  'Administrativo': '#8b5cf6',
  'TI': '#1e293b',
  'RegulaÃ§Ã£o': '#fb923c'
};

const ROOMS = ['Sala A', 'Sala B', 'Almoxarifado Principal', 'FarmÃ¡cia'];

const CATEGORY_COLORS: Record<string, string> = {
  'OdontolÃ³gico': '#0284c7',
  'MÃ©dico Hospitalar': '#ef4444',
  'AlimentÃ­cio': '#f59e0b',
  'Expediente': '#3b82f6',
  'Higiene': '#10b981',
  'RadiolÃ³gico': '#8b5cf6',
  'Saneante': '#06b6d4',
  'Copa & Cozinha': '#f97316',
  'Papelaria': '#0ea5e9',
  'EPI': '#ec4899',
  'GrÃ¡fica': '#fbbf24',
  'InformÃ¡tica': '#6366f1',
  'Limpeza': '#059669',
  'AnestÃ©sico': '#7c3aed',
  'Medicamentos': '#be123c',
  'FisioterÃ¡picos': '#14b8a6',
  'ManutenÃ§Ã£o': '#57534e',
  'Outros': '#78716c',
};

const getCategoryColor = (cat: string) => {
  if (CATEGORY_COLORS[cat]) return CATEGORY_COLORS[cat];
  const hash = cat.split('').reduce((acc, char) => char.charCodeAt(0) + ((acc << 5) - acc), 0);
  return `hsl(${Math.abs(hash) % 360}, 70%, 50%)`;
};

enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId: string | undefined;
    email: string | null | undefined;
    emailVerified: boolean | undefined;
    isAnonymous: boolean | undefined;
    tenantId: string | null | undefined;
    providerInfo: {
      providerId: string;
      displayName: string | null;
      email: string | null;
      photoUrl: string | null;
    }[];
  }
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errMessage = error instanceof Error ? error.message : String(error);
  const errInfo: FirestoreErrorInfo = {
    error: errMessage,
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      tenantId: auth.currentUser?.tenantId,
      providerInfo: auth.currentUser?.providerData?.map(provider => ({
        providerId: provider.providerId,
        displayName: provider.displayName,
        email: provider.email,
        photoUrl: provider.photoURL
      })) || []
    },
    operationType,
    path
  };

  console.error('Firestore Error: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

interface ErrorBoundaryProps {
  children: React.ReactNode;
}

interface ErrorBoundaryState {
  hasError: boolean;
  error: any;
}

export class ErrorBoundary extends React.Component<ErrorBoundaryProps, ErrorBoundaryState> {
  constructor(props: ErrorBoundaryProps) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error: any) {
    return { hasError: true, error };
  }

  componentDidCatch(error: any, errorInfo: any) {
    handleFirestoreError(error, OperationType.WRITE, 'client_crash');
  }

  render() {
    if (this.state.hasError) {
      const isQuotaError = String(this.state.error?.message || '').toLowerCase().includes('quota') || 
                           String(this.state.error?.message || '').toLowerCase().includes('resource_exhausted') ||
                           String(this.state.error?.message || '').toLowerCase().includes('resource-exhausted');

      return (
        <div className="min-h-screen bg-slate-900 flex items-center justify-center p-6">
          <div className="bg-white p-8 rounded-3xl shadow-2xl max-w-lg w-full border border-slate-100">
            <div className={`w-16 h-16 ${isQuotaError ? 'bg-amber-100 text-amber-600' : 'bg-rose-100 text-rose-600'} rounded-2xl flex items-center justify-center mx-auto mb-6`}>
              <AlertTriangle size={32} />
            </div>
            
            {isQuotaError ? (
              <>
                <h2 className="text-2xl font-black text-center text-slate-900 mb-2">Cota Gratuita de Leituras Excedida</h2>
                <p className="text-slate-600 text-xs font-medium text-center mb-6 leading-relaxed">
                  A cota diÃ¡ria do plano gratuito do Firebase (leituras de banco de dados) foi temporariamente atingida. 
                  O sistema continua armazenando suas alteraÃ§Ãµes e ativou a aceleraÃ§Ã£o por cache local no navegador.
                </p>
                <div className="bg-amber-50 border border-amber-200 p-4 rounded-2xl mb-6">
                  <p className="text-xs font-bold text-amber-800 mb-1"> O que vocÃª pode fazer:</p>
                  <ul className="text-[11px] text-amber-700 space-y-1 list-disc list-inside">
                    <li>Recarregue a pÃ¡gina para utilizar os dados em cache no seu navegador.</li>
                    <li>As cotas diÃ¡rias gratuitas sÃ£o renovadas automaticamente pelo Google Firebase a cada novo ciclo diÃ¡rio.</li>
                  </ul>
                </div>
              </>
            ) : (
              <>
                <h2 className="text-2xl font-black text-center text-slate-900 mb-4">Algo deu errado</h2>
                <p className="text-slate-500 text-center mb-6 text-sm">
                  Ocorreu um erro inesperado. Por favor, recarregue a pÃ¡gina ou tente novamente.
                </p>
                <div className="bg-rose-50 p-4 rounded-xl mb-6 overflow-auto max-h-40 border border-rose-100">
                  <code className="text-xs text-rose-700">
                    {this.state.error?.message || "Erro desconhecido"}
                  </code>
                </div>
              </>
            )}

            <button 
              onClick={() => window.location.reload()}
              className="w-full py-4 bg-gradient-to-r from-blue-700 to-indigo-900 text-white rounded-2xl font-black text-sm hover:from-blue-800 hover:to-indigo-950 transition-all shadow-lg shadow-blue-900/20"
            >
              Recarregar AplicaÃ§Ã£o
            </button>
          </div>
        </div>
      );
    }

    return this.props.children;
  }
}

export default function App() {
  const [items, setItems] = useState<Item[]>([]);
  const [transactions, setTransactions] = useState<Transaction[]>([]);
  const [requests, setRequests] = useState<MaterialRequest[]>([]);
  const [allRequestItems, setAllRequestItems] = useState<RequestItem[]>([]);
  const [user, setUser] = useState<User | null>(null);
  const [userProfile, setUserProfile] = useState<UserProfile | null>(null);
  const [loading, setLoading] = useState(true);
  const [loginLoading, setLoginLoading] = useState(false);
  const [isRegistering, setIsRegistering] = useState(false);
  const [authEmail, setAuthEmail] = useState('');
  const [authPassword, setAuthPassword] = useState('');
  const [authName, setAuthName] = useState('');
  const [authSectors, setAuthSectors] = useState<string[]>([]);
  const [selectedSector, setSelectedSector] = useState(SECTORS[0]);
  const [donationUnitName, setDonationUnitName] = useState('');
  const [donationUnitAddress, setDonationUnitAddress] = useState('');
  const [donationUnitCNPJ, setDonationUnitCNPJ] = useState('');
  const [donationRevisionDate, setDonationRevisionDate] = useState('');
  const [letterheadImage, setLetterheadImage] = useState<string | null>(null);
  const [reportsTab, setReportsTab] = useState<'overview' | 'quantitativo' | 'letterhead'>('overview');
  const [quantitativoSource, setQuantitativoSource] = useState<'sample' | 'system'>('system');
  const [quantitativoPeriodPreset, setQuantitativoPeriodPreset] = useState<'1_semestre_2026' | '2_semestre_2026' | 'ano_2026' | 'custom'>('1_semestre_2026');
  const [quantitativoCustomStart, setQuantitativoCustomStart] = useState('2026-01-01');
  const [quantitativoCustomEnd, setQuantitativoCustomEnd] = useState('2026-06-30');
  const [quantitativoCategory, setQuantitativoCategory] = useState('Material MÃ©dico-Hospitalar');
  const [quantitativoTitle, setQuantitativoTitle] = useState('');
  const [quantitativoCriticalAnalysis, setQuantitativoCriticalAnalysis] = useState('');
  const [isEditingQuantitativoAnalysis, setIsEditingQuantitativoAnalysis] = useState(false);
  const quantitativoReportRef = useRef<HTMLDivElement>(null);
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'inventory' | 'history' | 'requests' | 'admin-devolutions' | 'reports' | 'my-requests' | 'new-request' | 'devolution' | 'users' | 'trash' | 'leader-stats'>('dashboard');
  const leaderStatistics = useMemo(() => {
    if (userProfile?.role !== 'LÃDER' && userProfile?.role !== 'SETOR') return { topRequested: [], topDelivered: [] };

    const requestedMap: Record<string, number> = {};
    const deliveredMap: Record<string, number> = {};

    allRequestItems.forEach(item => {
      const parentRequest = requests.find(r => r.id === item.request_id);
      if (!parentRequest || parentRequest.sector !== selectedSector) return;

      const normalizedName = item.product_name;
      requestedMap[normalizedName] = (requestedMap[normalizedName] || 0) + (item.quantity_requested || 0);

      if (parentRequest.status === 'ENTREGUE') {
        deliveredMap[normalizedName] = (deliveredMap[normalizedName] || 0) + (item.quantity_approved || 0);
      }
    });

    const topRequested = Object.entries(requestedMap)
      .map(([name, qty]) => ({ name, qty }))
      .sort((a, b) => b.qty - a.qty)
      .slice(0, 10);

    const topDelivered = Object.entries(deliveredMap)
      .map(([name, qty]) => ({ name, qty }))
      .sort((a, b) => b.qty - a.qty)
      .slice(0, 10);

    return { topRequested, topDelivered };
  }, [allRequestItems, requests, userProfile, selectedSector]);
  const [showAddModal, setShowAddModal] = useState(false);
  const [showTransactionModal, setShowTransactionModal] = useState<{show: boolean, type: 'entry' | 'exit', item?: Item}>({ show: false, type: 'entry' });
  const [transactionMinStock, setTransactionMinStock] = useState<number>(NaN);
  const [showDetailModal, setShowDetailModal] = useState<{show: boolean, type: 'low_stock' | 'expiry' | 'all_alerts', items: (Item | ItemGroup)[]}>({ show: false, type: 'low_stock', items: [] });
  const [showDeleteModal, setShowDeleteModal] = useState<{show: boolean, transactionId?: string}>({ show: false });
  const [showSettingsModal, setShowSettingsModal] = useState(false);
  const [settingsTab, setSettingsTab] = useState<'logo' | 'tools' | 'info'>('logo');
  const [distribViewMode, setDistribViewMode] = useState<'types' | 'units'>('types');
  const [showMergeSuppliers, setShowMergeSuppliers] = useState(false);
  const [showMergeItems, setShowMergeItems] = useState(false);
  const [sourceSupplier, setSourceSupplier] = useState('');
  const [targetSupplier, setTargetSupplier] = useState('');
  const [sourceItemName, setSourceItemName] = useState('');
  const [targetItemName, setTargetItemName] = useState('');
  const [isMerging, setIsMerging] = useState(false);
  const [showUserDeleteConfirm, setShowUserDeleteConfirm] = useState<{show: boolean, user?: UserProfile}>({ show: false });
  const [showDeleteTestDataModal, setShowDeleteTestDataModal] = useState(false);
  const [deleteTestTarget, setDeleteTestTarget] = useState<'entries_only' | 'entries_and_stock' | 'all_test_data'>('entries_only');
  const [deleteTestConfirmInput, setDeleteTestConfirmInput] = useState('');
  const [isDeletingTestData, setIsDeletingTestData] = useState(false);
  const [toast, setToast] = useState<{show: boolean, message: string, type: 'success' | 'error' | 'info'}>({ show: false, message: '', type: 'info' });
  const [showRequestDetailModal, setShowRequestDetailModal] = useState<{show: boolean, request?: MaterialRequest}>({ show: false });
  const [showDevolutionModal, setShowDevolutionModal] = useState<{show: boolean, request?: MaterialRequest}>({ show: false });
  const [devolutionBasket, setDevolutionBasket] = useState<Array<{ product_id: string, product_name: string, quantity: number, maxQty: number, selectedBatchId: string }>>([]);
  const [selectedDevProduct, setSelectedDevProduct] = useState('');
  const [devolutionReason, setDevolutionReason] = useState('NÃ£o teve uso');
  const [devolutionObservation, setDevolutionObservation] = useState('');
  const [isProcessingDevolution, setIsProcessingDevolution] = useState(false);
  const [devolutionSubTab, setDevolutionSubTab] = useState<'my_returns' | 'eligible_deliveries' | 'sector_stock'>('my_returns');
  const [adminAddItemSearch, setAdminAddItemSearch] = useState('');
  const [isAdminAddingItem, setIsAdminAddingItem] = useState(false);
  const [notifications, setNotifications] = useState<Notification[]>([]);
  const [showNotifications, setShowNotifications] = useState(false);
  const [showStockConfirm, setShowStockConfirm] = useState<{show: boolean, notificationId?: string, itemName?: string}>({show: false});
  const [usersList, setUsersList] = useState<UserProfile[]>([]);
  const [deletionReason, setDeletionReason] = useState('');
  const [showDeletedHistory, setShowDeletedHistory] = useState(false);
  const [categoryFilter, setCategoryFilter] = useState('all');
  const [inventorySort, setInventorySort] = useState<'name_asc' | 'name_desc' | 'duration_asc' | 'duration_desc'>('name_asc');
  const [inventoryLocation, setInventoryLocation] = useState<'Almoxarifado' | 'FarmÃ¡cia'>('Almoxarifado');

  const isAdmin = userProfile?.role === 'ADMIN' || 
                  user?.email === 'gerlianemagalhaes79@gmail.com' || 
                  userProfile?.sector === 'Almoxarifado';

  useEffect(() => {
    if (userProfile?.sector === 'FarmÃ¡cia' || selectedSector === 'FarmÃ¡cia') {
      if (!isAdmin) {
        setInventoryLocation('FarmÃ¡cia');
      }
    } else if (!isAdmin) {
      setInventoryLocation('Almoxarifado');
    }
  }, [selectedSector, userProfile?.sector, isAdmin]);

  const weeklyExitRates = useMemo(() => {
    const twentyOneDaysAgo = new Date();
    twentyOneDaysAgo.setDate(twentyOneDaysAgo.getDate() - 21);
    
    const rates: Record<string, number> = {};
    
    transactions.forEach(t => {
      if (t.type === 'exit' && !t.deletedAt && new Date(t.date) >= twentyOneDaysAgo) {
        rates[t.item_name] = (rates[t.item_name] || 0) + t.quantity;
      }
    });
    
    // Convert to weekly average (21 days is exactly 3 weeks)
    Object.keys(rates).forEach(name => {
      rates[name] = rates[name] / 3;
    });
    
    return rates;
  }, [transactions]);

  // Request states
  const [requestBasket, setRequestBasket] = useState<{product_id: string, product_name: string, quantity: number}[]>([]);
  const [requestObservation, setRequestObservation] = useState('');
  const [adminObservation, setAdminObservation] = useState('');
  const [isSyncingStock, setIsSyncingStock] = useState(false);

  // Auto-update Minimum Stock based on consumption velocity (5 weeks coverage)
  useEffect(() => {
    if (!isAdmin || items.length === 0 || transactions.length === 0 || isSyncingStock) return;

    const syncStockVelocity = async () => {
      const updates: { id: string, newMin: number }[] = [];
      const now = new Date();
      
      // We only consider items with enough history (e.g., at least 1 exit in the last 21 days)
      Object.keys(weeklyExitRates).forEach(itemName => {
        const weeklyRate = weeklyExitRates[itemName];
        if (weeklyRate > 0) {
          const recommendedMin = Math.ceil(weeklyRate * 5);
          
          // Find all batches of this item and check if their min_quantity needs update
          items.forEach(item => {
            if (item.name === itemName && !item.deletedAt) {
              // Only update if difference is more than 0 and actually different from stored
              if (recommendedMin !== item.min_quantity) {
                updates.push({ id: item.id, newMin: recommendedMin });
              }
            }
          });
        }
      });

      if (updates.length > 0) {
        setIsSyncingStock(true);
        try {
          console.log(`Auto-otimizando estoque mÃ­nimo para ${updates.length} lotes...`);
          // Batch updates to Firestore (max 500 per batch)
          for (let i = 0; i < updates.length; i += 450) {
            const batch = writeBatch(db);
            const chunk = updates.slice(i, i + 450);
            chunk.forEach(u => {
              batch.update(doc(db, 'items', u.id), {
                min_quantity: u.newMin,
                updatedAt: serverTimestamp()
              });
            });
            await batch.commit();
          }
          console.log("OtimizaÃ§Ã£o de estoque mÃ­nimo concluÃ­da.");
        } catch (error) {
          console.error("Erro ao auto-atualizar estoques mÃ­nimos:", error);
        } finally {
          setIsSyncingStock(false);
        }
      }
    };

    // Run sync after a short delay once data is loaded, and then every hour if the tab stays open
    const initialSync = setTimeout(syncStockVelocity, 10000);
    const intervalSync = setInterval(syncStockVelocity, 3600000); 

    return () => {
      clearTimeout(initialSync);
      clearInterval(intervalSync);
    };
  }, [isAdmin, items, transactions, weeklyExitRates, isSyncingStock]);
  const [isSubmittingRequest, setIsSubmittingRequest] = useState(false);
  const [editingRequest, setEditingRequest] = useState<MaterialRequest | null>(null);
  const [showRoomInventoryModal, setShowRoomInventoryModal] = useState(false);
  const [selectedRoom, setSelectedRoom] = useState('Sala A');
  const [customRoomName, setCustomRoomName] = useState('Sala A');
  const [selectedRoomCategories, setSelectedRoomCategories] = useState<string[]>([]);
  
  const createNotification = async (userId: string, title: string, message: string, requestId?: string, type: 'STOCK_ZERO' | 'SYSTEM' | 'REQUEST' = 'SYSTEM', itemName?: string) => {
    try {
      const data: any = {
        userId,
        title,
        message,
        date: new Date().toISOString(),
        read: false,
        type,
      };
      
      if (requestId !== undefined) {
        data.requestId = requestId;
      }
      if (itemName !== undefined) {
        data.itemName = itemName;
      }

      await addDoc(collection(db, 'notifications'), data);
    } catch (error) {
      console.error("Error creating notification:", error);
    }
  };

  const checkStockAndNotify = async (itemName: string) => {
    try {
      if (!itemName) return;
      
      const normalizedName = normalizeString(itemName);
      const safeId = getSafeDocId(itemName);
      
      // Get all active batches for this product by scanning items and filtering in memory
      const itemsSnapshot = await getDocs(collection(db, 'items'));
      const batches = itemsSnapshot.docs
        .map(d => ({ id: d.id, ...d.data() } as Item))
        .filter(i => !i.deletedAt && normalizeString(i.name) === normalizedName);

      const totalQuantity = batches.reduce((sum, b) => sum + (Number(b.quantity) || 0), 0);

      if (totalQuantity === 0) {
        // Check if user already acknowledged this zero stock alert
        const dismissalDoc = await getDoc(doc(db, 'dismissed_stock_alerts', safeId));
        if (dismissalDoc.exists()) {
          return;
        }

        // Check if an unconfirmed notification for this item already exists
        const existingNotifQuery = query(
          collection(db, 'notifications'),
          where('userId', '==', 'ADMIN_GROUP'),
          where('read', '==', false)
        );
        const existingSnap = await getDocs(existingNotifQuery);
        
        // In-memory robust check to cover casing & spacing variations
        const alreadyNotified = existingSnap.docs.some(d => {
          const data = d.data();
          return data.type === 'STOCK_ZERO' && normalizeString(data.itemName) === normalizedName;
        });

        if (!alreadyNotified) {
          await createNotification(
            'ADMIN_GROUP',
            'Estoque Zerado',
            `O material "${itemName.toUpperCase()}" atingiu estoque zero.`,
            undefined,
            'STOCK_ZERO',
            itemName.toUpperCase()
          );
        }
      } else {
        // If stock goes back up, clear the dismissal entry to enable future zero alerts
        try {
          await deleteDoc(doc(db, 'dismissed_stock_alerts', safeId));
        } catch (e) {
          // Ignore
        }
      }
    } catch (error) {
      console.error("Error checking stock for notification:", error);
    }
  };
  
  // Form states
  const [bulkEntry, setBulkEntry] = useState({
    supplier: '',
    category: 'Expediente',
    origin: 'extra' as 'contract' | 'extra' | 'donation',
    room: 'Almoxarifado Principal',
    items: [{
      id: Math.random().toString(36).substr(2, 9),
      name: '',
      initial_quantity: 1,
      min_quantity: NaN,
      batch_number: '',
      expiry_date: '',
      is_indeterminate_expiry: false,
      unit_price: 0,
      unit_measure: 'Unidade (UN)',
      medication_type: ''
    }]
  });
  const [categories, setCategories] = useState<string[]>(['OdontolÃ³gico', 'MÃ©dico Hospitalar', 'AlimentÃ­cio', 'Expediente', 'Higiene', 'RadiolÃ³gico', 'Saneante', 'Copa & Cozinha', 'Papelaria', 'EPI', 'GrÃ¡fica', 'InformÃ¡tica', 'Limpeza', 'AnestÃ©sico', 'Medicamentos', 'FisioterÃ¡picos', 'ManutenÃ§Ã£o']);
  const [showNewCategoryInput, setShowNewCategoryInput] = useState(false);
  const [newCategoryName, setNewCategoryName] = useState('');

  useEffect(() => {
    if (showAddModal) {
      setBulkEntry({
        supplier: '',
        category: inventoryLocation === 'FarmÃ¡cia' ? 'Medicamentos' : 'Expediente',
        origin: 'extra' as 'contract' | 'extra' | 'donation',
        room: inventoryLocation === 'FarmÃ¡cia' ? 'FarmÃ¡cia' : 'Almoxarifado Principal',
        items: [{
          id: Math.random().toString(36).substr(2, 9),
          name: '',
          initial_quantity: 1,
          min_quantity: NaN,
          batch_number: '',
          expiry_date: '',
          is_indeterminate_expiry: false,
          unit_price: 0,
          unit_measure: 'Unidade (UN)',
          medication_type: ''
        }]
      });
    }
  }, [showAddModal, inventoryLocation]);

  const addBulkItemRow = () => {
    setBulkEntry(prev => ({
      ...prev,
      items: [...prev.items, {
        id: Math.random().toString(36).substr(2, 9),
        name: '',
        initial_quantity: 1,
        min_quantity: NaN,
        batch_number: '',
        expiry_date: '',
        is_indeterminate_expiry: false,
        unit_price: 0,
        unit_measure: 'Unidade (UN)',
        medication_type: ''
      }]
    }));
  };

  const removeBulkItemRow = (id: string) => {
    if (bulkEntry.items.length > 1) {
      setBulkEntry(prev => ({
        ...prev,
        items: prev.items.filter(item => item.id !== id)
      }));
    }
  };

  const duplicateBulkItem = (id: string) => {
    const itemToDuplicate = bulkEntry.items.find(item => item.id === id);
    if (itemToDuplicate) {
      setBulkEntry(prev => ({
        ...prev,
        items: [...prev.items, {
          ...itemToDuplicate,
          id: Math.random().toString(36).substr(2, 9),
          batch_number: '',
          initial_quantity: 1,
          expiry_date: '',
          unit_measure: itemToDuplicate.unit_measure || 'Unidade (UN)',
          medication_type: itemToDuplicate.medication_type || ''
        }]
      }));
    }
  };

  const updateBulkItem = (id: string, field: string, value: any) => {
    setBulkEntry(prev => ({
      ...prev,
      items: prev.items.map(item => {
        if (item.id === id) {
          let processedValue = value;
          if (field === 'name' && typeof value === 'string') {
            processedValue = value.toUpperCase();
          }
          const updatedItem = { ...item, [field]: processedValue };
          
          // Auto-fill min_quantity if name is changed and we have a calculated rate
          if (field === 'name' && processedValue) {
            const weeklyRate = weeklyExitRates[processedValue] || 0;
            if (weeklyRate > 0) {
              updatedItem.min_quantity = Math.ceil(weeklyRate * 5);
            } else {
              // Try to find if the item exists but has no history yet, use its current min_quantity
              const existingItem = items.find(i => i.name === processedValue);
              if (existingItem) {
                updatedItem.min_quantity = existingItem.min_quantity;
              }
            }
          }
          
          return updatedItem;
        }
        return item;
      })
    }));
  };
  
  const [modalSector, setModalSector] = useState<string>('');
  const [transactionQty, setTransactionQty] = useState(1);
  const [exitReason, setExitReason] = useState<'consumo' | 'doacao' | 'vencido' | 'perda'>('consumo');
  const [expiryReason, setExpiryReason] = useState('');
  const [selectedItemId, setSelectedItemId] = useState<string>('');
  const [selectedItemName, setSelectedItemName] = useState<string>('');
  const [basket, setBasket] = useState<{item_id: string, quantity: number}[]>([]);
  const [searchTerm, setSearchTerm] = useState('');
  const [modalSearchTerm, setModalSearchTerm] = useState('');
  const [requestSearchTerm, setRequestSearchTerm] = useState('');
  const [reportRange, setReportRange] = useState({
    start: format(subDays(new Date(), 30), 'yyyy-MM-dd'),
    end: format(new Date(), 'yyyy-MM-dd')
  });
  const [printRange, setPrintRange] = useState({
    start: format(new Date(), 'yyyy-MM-dd'),
    end: format(new Date(), 'yyyy-MM-dd')
  });
  const [reportSectorFilter, setReportSectorFilter] = useState<string>('all');
  const [pcaRange, setPcaRange] = useState({
    start: format(subDays(new Date(), 365), 'yyyy-MM-dd'),
    end: format(new Date(), 'yyyy-MM-dd')
  });
  const [pcaCategory, setPcaCategory] = useState('all');
  const [originFilter, setOriginFilter] = useState<'all' | 'contract' | 'extra' | 'donation'>('all');

  const [expandedItems, setExpandedItems] = useState<Set<string>>(new Set());
  const [editingPrice, setEditingPrice] = useState<{ id: string, price: number } | null>(null);
  const [editingQuantity, setEditingQuantity] = useState<{ id: string, quantity: number } | null>(null);
  const [editingMaterialName, setEditingMaterialName] = useState<{ oldName: string, newName: string } | null>(null);
  const [editingCategory, setEditingCategory] = useState<{ name: string, currentCategory: string, itemId?: string } | null>(null);
  const [customNewCategory, setCustomNewCategory] = useState('');
  const [showChangeCategoryModal, setShowChangeCategoryModal] = useState(false);
  const [categoryModalMaterial, setCategoryModalMaterial] = useState('');
  const [categoryModalNewCategory, setCategoryModalNewCategory] = useState('');
  const [customModalCategory, setCustomModalCategory] = useState('');
  const [isUpdatingCategory, setIsUpdatingCategory] = useState(false);

  const uniqueSuppliers = useMemo(() => {
    const fromItems = items.map(i => i.supplier).filter(Boolean) as string[];
    const fromTrans = transactions.map(t => t.supplier).filter(Boolean) as string[];
    return Array.from(new Set([...fromItems, ...fromTrans])).sort();
  }, [items, transactions]);

  const uniqueItemNames = useMemo(() => {
    const names = new Set(items.filter(i => !i.deletedAt).map(i => i.name));
    return Array.from(names).sort((a, b) => a.localeCompare(b));
  }, [items]);

  useEffect(() => {
    if (showRequestDetailModal.show && showRequestDetailModal.request) {
      setAdminObservation(showRequestDetailModal.request.adminObservation || '');
    } else {
      setAdminObservation('');
    }
  }, [showRequestDetailModal.show, showRequestDetailModal.request]);

  const toggleExpand = (name: string) => {
    const newExpanded = new Set(expandedItems);
    if (newExpanded.has(name)) {
      newExpanded.delete(name);
    } else {
      newExpanded.add(name);
    }
    setExpandedItems(newExpanded);
  };

  useEffect(() => {
    if (showTransactionModal.show) {
      if (showTransactionModal.type === 'exit' && showTransactionModal.item) {
        setBasket([{ item_id: showTransactionModal.item.id, quantity: 1 }]);
      }
    } else {
      setModalSearchTerm('');
      setSelectedItemName('');
      setSelectedItemId('');
      if (showTransactionModal.type === 'exit') {
        setBasket([]);
      }
    }
  }, [showTransactionModal.show, showTransactionModal.type, showTransactionModal.item]);

  useEffect(() => {
    if (activeTab !== 'new-request') {
      setRequestSearchTerm('');
    }
  }, [activeTab]);

  const handleUpdatePrice = async () => {
    if (!editingPrice) return;
    try {
      const itemToUpdate = items.find(i => i.id === editingPrice.id);
      if (!itemToUpdate) return;

      // Update all items with the same name to keep prices consistent across batches
      const itemsWithSameName = items.filter(i => i.name.toLowerCase() === itemToUpdate.name.toLowerCase() && !i.deletedAt);
      
      const batch = writeBatch(db);
      itemsWithSameName.forEach(item => {
        batch.update(doc(db, 'items', item.id), {
          unit_price: editingPrice.price,
          updatedAt: serverTimestamp()
        });
      });
      
      await batch.commit();
      showToast(`PreÃ§o unitÃ¡rio de "${itemToUpdate.name}" atualizado em todos os lotes!`, "success");
      setEditingPrice(null);
    } catch (error: any) {
      handleFirestoreError(error, OperationType.UPDATE, `items/${editingPrice.id}`);
      showToast(`Erro ao atualizar preÃ§o: ${error.message}`, "error");
    }
  };

  const handleUpdateQuantity = async () => {
    if (!editingQuantity) return;
    try {
      await updateDoc(doc(db, 'items', editingQuantity.id), {
        quantity: editingQuantity.quantity
      });
      showToast("Quantidade atualizada com sucesso!", "success");
      const name = items.find(i => i.id === editingQuantity.id)?.name;
      setEditingQuantity(null);
      if (name) await checkStockAndNotify(name);
    } catch (error: any) {
      handleFirestoreError(error, OperationType.UPDATE, `items/${editingQuantity.id}`);
      showToast(`Erro ao atualizar quantidade: ${error.message}`, "error");
    }
  };

  const handleUpdateMaterialName = async () => {
    if (!editingMaterialName || !editingMaterialName.newName.trim()) return;
    const oldName = editingMaterialName.oldName;
    const newName = editingMaterialName.newName.trim();

    if (oldName === newName) {
      setEditingMaterialName(null);
      return;
    }

    try {
      // Find all items and transactions with the old name
      const itemsToUpdate = items.filter(i => i.name === oldName);
      const transToUpdate = transactions.filter(t => t.item_name === oldName);
      
      const totalOps = itemsToUpdate.length + transToUpdate.length;
      
      if (totalOps === 0) {
        setEditingMaterialName(null);
        return;
      }

      // Process in batches of 400
      const allDocs = [
        ...itemsToUpdate.map(i => ({ ref: doc(db, 'items', i.id), data: { name: newName } })),
        ...transToUpdate.map(t => ({ ref: doc(db, 'transactions', t.id), data: { item_name: newName } }))
      ];

      for (let i = 0; i < allDocs.length; i += 400) {
        const batch = writeBatch(db);
        const chunk = allDocs.slice(i, i + 400);
        chunk.forEach(op => batch.update(op.ref, op.data));
        await batch.commit();
      }

      showToast("Nome do material atualizado com sucesso!", "success");
      setEditingMaterialName(null);
    } catch (error: any) {
      console.error("Error updating material name:", error);
      showToast(`Erro ao atualizar nome: ${error.message}`, "error");
    }
  };

  const handleUpdateCategory = async (targetCategory?: string) => {
    if (!editingCategory) return;
    
    let newCat = (targetCategory !== undefined ? targetCategory : editingCategory.currentCategory).trim();
    if (newCat === '__NEW__') {
      newCat = customNewCategory.trim();
    }
    
    if (!newCat) {
      showToast("Por favor, selecione ou informe uma categoria vÃ¡lida.", "error");
      return;
    }

    try {
      if (editingCategory.itemId) {
        await updateDoc(doc(db, 'items', editingCategory.itemId), {
          category: newCat
        });
        showToast("Categoria do lote atualizada com sucesso!", "success");
      } else {
        const itemsToUpdate = items.filter(i => i.name === editingCategory.name);
        if (itemsToUpdate.length === 0) {
          setEditingCategory(null);
          return;
        }

        for (let i = 0; i < itemsToUpdate.length; i += 400) {
          const batch = writeBatch(db);
          const chunk = itemsToUpdate.slice(i, i + 400);
          chunk.forEach(item => {
            batch.update(doc(db, 'items', item.id), { category: newCat });
          });
          await batch.commit();
        }
        showToast(`Categoria de "${editingCategory.name}" alterada para "${newCat}" com sucesso!`, "success");
      }

      if (!categories.includes(newCat)) {
        setCategories(prev => [...prev, newCat]);
      }

      setEditingCategory(null);
      setCustomNewCategory('');
    } catch (error: any) {
      console.error("Error updating category:", error);
      handleFirestoreError(error, OperationType.UPDATE, `items`);
      showToast(`Erro ao atualizar categoria: ${error.message}`, "error");
    }
  };

  const handleModalChangeCategory = async () => {
    if (!categoryModalMaterial) {
      showToast("Selecione o material que deseja alterar.", "error");
      return;
    }

    let newCat = categoryModalNewCategory.trim();
    if (newCat === '__NEW__') {
      newCat = customModalCategory.trim();
    }

    if (!newCat) {
      showToast("Informe a nova categoria.", "error");
      return;
    }

    setIsUpdatingCategory(true);
    try {
      const itemsToUpdate = items.filter(i => i.name === categoryModalMaterial);
      if (itemsToUpdate.length === 0) {
        showToast("Nenhum item encontrado com esse nome.", "error");
        setIsUpdatingCategory(false);
        return;
      }

      for (let i = 0; i < itemsToUpdate.length; i += 400) {
        const batch = writeBatch(db);
        const chunk = itemsToUpdate.slice(i, i + 400);
        chunk.forEach(item => {
          batch.update(doc(db, 'items', item.id), { category: newCat });
        });
        await batch.commit();
      }

      if (!categories.includes(newCat)) {
        setCategories(prev => [...prev, newCat]);
      }

      showToast(`Categoria do material "${categoryModalMaterial}" alterada para "${newCat}" com sucesso!`, "success");
      setShowChangeCategoryModal(false);
      setCategoryModalMaterial('');
      setCategoryModalNewCategory('');
      setCustomModalCategory('');
    } catch (error: any) {
      console.error("Error changing category:", error);
      handleFirestoreError(error, OperationType.UPDATE, `items`);
      showToast(`Erro ao alterar categoria: ${error.message}`, "error");
    } finally {
      setIsUpdatingCategory(false);
    }
  };

  const handleMergeSuppliers = async () => {
    if (!sourceSupplier || !targetSupplier || sourceSupplier === targetSupplier) {
      showToast("Selecione fornecedores diferentes para mesclar.", "error");
      return;
    }

    setIsMerging(true);
    try {
      // Find all items and transactions with the source supplier
      const itemsToUpdate = items.filter(i => i.supplier === sourceSupplier);
      const transToUpdate = transactions.filter(t => t.supplier === sourceSupplier);
      
      const totalOps = itemsToUpdate.length + transToUpdate.length;
      
      if (totalOps === 0) {
        showToast("Nenhum registro encontrado para o fornecedor de origem.", "info");
        setIsMerging(false);
        return;
      }

      // Process in batches of 400 (Firestore limit is 500)
      const allDocs = [
        ...itemsToUpdate.map(i => ({ ref: doc(db, 'items', i.id), data: { supplier: targetSupplier } })),
        ...transToUpdate.map(t => ({ ref: doc(db, 'transactions', t.id), data: { supplier: targetSupplier } }))
      ];

      for (let i = 0; i < allDocs.length; i += 400) {
        const batch = writeBatch(db);
        const chunk = allDocs.slice(i, i + 400);
        chunk.forEach(op => batch.update(op.ref, op.data));
        await batch.commit();
      }

      showToast(`${totalOps} registros atualizados com sucesso!`, "success");
      setShowMergeSuppliers(false);
      setSourceSupplier('');
      setTargetSupplier('');
    } catch (error: any) {
      console.error("Error merging suppliers:", error);
      showToast(`Erro ao mesclar fornecedores: ${error.message}`, "error");
    } finally {
      setIsMerging(false);
    }
  };

  const handleMergeItems = async () => {
    if (!sourceItemName || !targetItemName || sourceItemName === targetItemName) {
      showToast("Selecione itens diferentes para mesclar.", "error");
      return;
    }

    setIsMerging(true);
    try {
      // Find all items and transactions with the source item name
      const itemsToUpdate = items.filter(i => i.name === sourceItemName);
      const transToUpdate = transactions.filter(t => t.item_name === sourceItemName);
      
      const totalOps = itemsToUpdate.length + transToUpdate.length;
      
      if (totalOps === 0) {
        showToast("Nenhum registro encontrado para o item de origem.", "info");
        setIsMerging(false);
        return;
      }

      // Process in batches of 400 (Firestore limit is 500)
      const allDocs = [
        ...itemsToUpdate.map(i => ({ ref: doc(db, 'items', i.id), data: { name: targetItemName } })),
        ...transToUpdate.map(t => ({ ref: doc(db, 'transactions', t.id), data: { item_name: targetItemName } }))
      ];

      for (let i = 0; i < allDocs.length; i += 400) {
        const batch = writeBatch(db);
        const chunk = allDocs.slice(i, i + 400);
        chunk.forEach(op => batch.update(op.ref, op.data));
        await batch.commit();
      }

      showToast(`${totalOps} registros atualizados com sucesso!`, "success");
      setShowMergeItems(false);
      setSourceItemName('');
      setTargetItemName('');
    } catch (error: any) {
      console.error("Error merging items:", error);
      showToast(`Erro ao mesclar itens: ${error.message}`, "error");
    } finally {
      setIsMerging(false);
    }
  };

  const handleDeleteTestData = async () => {
    const confirmation = deleteTestConfirmInput.trim().toUpperCase();
    if (confirmation !== 'TESTE' && confirmation !== 'CONFIRMAR' && confirmation !== 'EXCLUIR') {
      showToast("Digite CONFIRMAR, TESTE ou EXCLUIR para autorizar a remoÃ§Ã£o.", "error");
      return;
    }

    setIsDeletingTestData(true);
    try {
      if (deleteTestTarget === 'entries_only') {
        const entryTrans = transactions.filter(t => t.type === 'entry');
        if (entryTrans.length === 0) {
          showToast("Nenhuma entrada de material encontrada para excluir.", "info");
          setIsDeletingTestData(false);
          return;
        }

        // Delete all entry transactions
        const docsToDelete = entryTrans.map(t => doc(db, 'transactions', t.id));
        for (let i = 0; i < docsToDelete.length; i += 400) {
          const batch = writeBatch(db);
          docsToDelete.slice(i, i + 400).forEach(d => batch.delete(d));
          await batch.commit();
        }

        // Update items stock balance to 0 and clear batches
        const itemUpdates = items.map(item => ({
          ref: doc(db, 'items', item.id),
          data: {
            quantity: 0,
            batches: [],
            updatedAt: serverTimestamp()
          }
        }));

        for (let i = 0; i < itemUpdates.length; i += 400) {
          const batch = writeBatch(db);
          itemUpdates.slice(i, i + 400).forEach(u => batch.update(u.ref, u.data));
          await batch.commit();
        }

        showToast(`${entryTrans.length} entradas de materiais de teste foram excluÃ­das com sucesso!`, "success");
      } else if (deleteTestTarget === 'entries_and_stock') {
        const entryTrans = transactions.filter(t => t.type === 'entry');
        const docsToDelete = [
          ...entryTrans.map(t => doc(db, 'transactions', t.id)),
          ...items.map(i => doc(db, 'items', i.id))
        ];

        for (let i = 0; i < docsToDelete.length; i += 400) {
          const batch = writeBatch(db);
          docsToDelete.slice(i, i + 400).forEach(d => batch.delete(d));
          await batch.commit();
        }

        showToast(`Todas as entradas e itens de estoque cadastrados como teste foram excluÃ­dos!`, "success");
      } else if (deleteTestTarget === 'all_test_data') {
        const reqItemsSnap = await getDocs(collection(db, 'request_items'));
        const docsToDelete = [
          ...transactions.map(t => doc(db, 'transactions', t.id)),
          ...items.map(i => doc(db, 'items', i.id)),
          ...requests.map(r => doc(db, 'requests', r.id)),
          ...reqItemsSnap.docs.map(d => d.ref)
        ];

        for (let i = 0; i < docsToDelete.length; i += 400) {
          const batch = writeBatch(db);
          docsToDelete.slice(i, i + 400).forEach(d => batch.delete(d));
          await batch.commit();
        }

        showToast(`Todos os dados de teste (entradas, saÃ­das, requisiÃ§Ãµes e estoque) foram completamente limpos!`, "success");
      }

      setShowDeleteTestDataModal(false);
      setDeleteTestConfirmInput('');
    } catch (error: any) {
      console.error("Error deleting test data:", error);
      showToast(`Erro ao excluir dados de teste: ${error.message}`, "error");
    } finally {
      setIsDeletingTestData(false);
    }
  };

  const showToast = (message: string, type: 'success' | 'error' | 'info' = 'info') => {
    setToast({ show: true, message, type });
    setTimeout(() => setToast({ show: false, message: '', type: 'info' }), 4000);
  };

  useEffect(() => {
    const unsubscribeAuth = onAuthStateChanged(auth, async (user) => {
      try {
        setUser(user);
        if (user) {
          const userEmail = user.email?.toLowerCase().trim();
          if (!userEmail) {
            await signOut(auth);
            showToast("Erro: E-mail nÃ£o encontrado no login do Google.", "error");
            setLoading(false);
            return;
          }

          // Always use email as the document ID for consistency
          const userRef = doc(db, 'users', userEmail);
          let userSnap: any = null;
          try {
            userSnap = await getDoc(userRef);
          } catch (e: any) {
            console.warn("Could not fetch user profile from Firestore:", e?.message || e);
          }

          if (userSnap && !userSnap.exists() && userEmail === 'gerlianemagalhaes79@gmail.com') {
            try {
              await setDoc(userRef, {
                email: userEmail,
                name: user.displayName || 'Admin',
                role: 'ADMIN',
                sector: 'Almoxarifado',
                uid: user.uid,
                lastLogin: new Date().toISOString()
              });
            } catch (e) {
              console.warn("Could not set master admin profile doc:", e);
            }
          } else if (userSnap && userSnap.exists()) {
            // Update existing profile with UID and last login
            try {
              await updateDoc(userRef, { 
                uid: user.uid,
                lastLogin: new Date().toISOString() 
              });
            } catch (e) {
              console.warn("Could not update last login timestamp:", e);
            }
          } else if (!userSnap && userEmail === 'gerlianemagalhaes79@gmail.com') {
            // Fallback for master admins when quota limit is exceeded
            setUserProfile({
              id: userEmail,
              name: user.displayName || 'Admin',
              role: 'ADMIN',
              sector: 'Almoxarifado',
              email: userEmail
            });
            setActiveTab('dashboard');
          } else if (userSnap && !userSnap.exists()) {
            // Not pre-registered and not master admin
            await signOut(auth);
            showToast("Acesso negado: Seu e-mail nÃ£o estÃ¡ cadastrado no sistema. Entre em contato com o administrador.", "error");
            setLoading(false);
            return;
          }

          onSnapshot(userRef, (doc) => {
            if (doc.exists()) {
              const profile = { id: doc.id, ...doc.data() } as UserProfile;
              setUserProfile(profile);
              
              if (profile.allowedSectors && profile.allowedSectors.length > 0) {
                setSelectedSector(prev => (prev && profile.allowedSectors?.includes(prev) ? prev : profile.allowedSectors![0]));
              } else if (profile.sector) {
                setSelectedSector(profile.sector);
              }

              if (profile.role === 'ADMIN' || userEmail === 'gerlianemagalhaes79@gmail.com' || profile.sector === 'Almoxarifado') {
                setActiveTab('dashboard');
              } else {
                setActiveTab('my-requests');
              }
            }
          }, (error) => {
            handleFirestoreError(error, OperationType.GET, `users/${userEmail}`);
          });
        } else {
          setUserProfile(null);
        }
      } catch (error: any) {
        const errStr = String(error?.message || error);
        if (errStr.toLowerCase().includes('quota limit exceeded') || errStr.toLowerCase().includes('resource_exhausted')) {
          console.warn("Auth state change notice (quota limit):", errStr);
        } else {
          console.error("Auth state change error:", error);
          showToast(`Erro na autenticaÃ§Ã£o: ${error.message}`, "error");
        }
      } finally {
        setLoading(false);
      }
    });
    return () => unsubscribeAuth();
  }, []);

  useEffect(() => {
    if (!user) {
      setItems([]);
      setTransactions([]);
      return;
    }

    const qItems = query(collection(db, 'items'), orderBy('name', 'asc'));
    const unsubscribeItems = onSnapshot(qItems, (snapshot) => {
      const itemsData = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as Item));
      setItems(itemsData);
      
      // Update categories list from existing items
      const existingCategories = Array.from(new Set(itemsData.map(i => i.category).filter(Boolean))) as string[];
      setCategories(prev => Array.from(new Set([...prev, ...existingCategories])));
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'items');
    });

    const qTrans = query(collection(db, 'transactions'), orderBy('date', 'desc'));
    const unsubscribeTrans = onSnapshot(qTrans, (snapshot) => {
      const fifteenDaysAgo = subDays(new Date(), 15);
      const transData = snapshot.docs
        .map(doc => ({ id: doc.id, ...doc.data() } as Transaction))
        .filter(t => !t.deletedAt || new Date(t.deletedAt) > fifteenDaysAgo);
      setTransactions(transData);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'transactions');
    });

    const qRequests = query(collection(db, 'requests'), orderBy('date', 'desc'));
    const unsubscribeRequests = onSnapshot(qRequests, (snapshot) => {
      const requestsData = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as MaterialRequest));
      setRequests(requestsData);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'requests');
    });

    const qReqItems = query(collection(db, 'request_items'));
    const unsubscribeReqItems = onSnapshot(qReqItems, (snapshot) => {
      const itemsData = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as RequestItem));
      setAllRequestItems(itemsData);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'request_items');
    });

    return () => {
      unsubscribeItems();
      unsubscribeTrans();
      unsubscribeRequests();
      unsubscribeReqItems();
    };
  }, [user]);

  useEffect(() => {
    if (!user) {
      setNotifications([]);
      return;
    }

    const notificationIds = [user.uid];
    if (isAdmin) {
      notificationIds.push('ADMIN_GROUP');
    }

    const qNotifications = query(
      collection(db, 'notifications'), 
      where('userId', 'in', notificationIds), 
      orderBy('date', 'desc')
    );
    
    const unsubscribeNotifications = onSnapshot(qNotifications, (snapshot) => {
      setNotifications(snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as Notification)));
    }, (error) => {
      // If error is permission denied, it might be because we added ADMIN_GROUP wrongly or rules haven't propagated
      console.warn("Notification listener error:", error);
      // Fallback to single user listener if needed
      const fallbackQ = query(collection(db, 'notifications'), where('userId', '==', user.uid), orderBy('date', 'desc'));
      const unsubFallback = onSnapshot(fallbackQ, (snapshot) => {
        setNotifications(snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as Notification)));
      });
      return () => unsubFallback();
    });

    return () => {
      unsubscribeNotifications();
    };
  }, [user, isAdmin]);

  // Retroactive check to sync materials that are currently with zero stock as notifications for administrators
  useEffect(() => {
    if (!isAdmin || items.length === 0) return;

    const runRetroactiveStockCheck = async () => {
      try {
        const activeItems = items.filter(i => !i.deletedAt);
        const groupedByName: { [key: string]: { name: string, totalQty: number } } = {};
        
        activeItems.forEach(item => {
          const normName = normalizeString(item.name);
          if (!groupedByName[normName]) {
            groupedByName[normName] = { name: item.name, totalQty: 0 };
          }
          groupedByName[normName].totalQty += (Number(item.quantity) || 0);
        });

        // Fetch currently dismissed alerts once to perform in-memory checks
        const dismissedSnap = await getDocs(collection(db, 'dismissed_stock_alerts'));
        const dismissedMap = new Set(dismissedSnap.docs.map(d => d.id));

        // 1. Clean up dismissal records for items that now have stock > 0
        const activeGrouped = Object.values(groupedByName);
        const withStock = activeGrouped.filter(g => g.totalQty > 0);
        for (const itemWithStock of withStock) {
          const safeId = getSafeDocId(itemWithStock.name);
          if (dismissedMap.has(safeId)) {
            try {
              await deleteDoc(doc(db, 'dismissed_stock_alerts', safeId));
              dismissedMap.delete(safeId);
            } catch (e) {
              // Ignore if doc doesn't exist or deletion fails
            }
          }
        }

        // 2. Identify and notify about zero stock items that haven't been dismissed or notified yet
        const zeroStockItems = activeGrouped.filter(g => g.totalQty === 0);
        const unreadStockZeroNotifications = notifications.filter(n => !n.read && n.type === 'STOCK_ZERO');

        for (const zeroItem of zeroStockItems) {
          const normZeroName = normalizeString(zeroItem.name);
          const safeId = getSafeDocId(zeroItem.name);
          const alreadyNotified = unreadStockZeroNotifications.some(n => normalizeString(n.itemName) === normZeroName);

          if (!alreadyNotified) {
            // Check if administrators have already confirmed science for this zero stock event
            if (dismissedMap.has(safeId)) {
              continue;
            }

            const existingNotifQuery = query(
              collection(db, 'notifications'),
              where('userId', '==', 'ADMIN_GROUP'),
              where('read', '==', false)
            );
            const existingSnap = await getDocs(existingNotifQuery);
            const alreadyInFirestore = existingSnap.docs.some(d => {
              const data = d.data();
              return data.type === 'STOCK_ZERO' && normalizeString(data.itemName) === normZeroName;
            });

            if (!alreadyInFirestore) {
              await createNotification(
                'ADMIN_GROUP',
                'Estoque Zerado',
                `O material "${zeroItem.name.toUpperCase()}" atingiu estoque zero.`,
                undefined,
                'STOCK_ZERO',
                zeroItem.name.toUpperCase()
              );
            }
          }
        }
      } catch (err) {
        console.warn("Notice in retroactive zero stock synchronization:", err);
      }
    };

    const timer = setTimeout(() => {
      runRetroactiveStockCheck();
    }, 2500);

    return () => clearTimeout(timer);
  }, [items, notifications, isAdmin]);

  useEffect(() => {
    if (!user || !userProfile) return;
    
    let unsubscribeUsers = () => {};
    if (user.email === 'gerlianemagalhaes79@gmail.com' || userProfile.role === 'ADMIN' || selectedSector === 'Almoxarifado') {
      // Ensure master admin exists in the database
      const masterAdmins = [
        { email: 'gerlianemagalhaes79@gmail.com', name: 'Admin' }
      ];

      masterAdmins.forEach(async (admin) => {
        const adminRef = doc(db, 'users', admin.email);
        const adminSnap = await getDoc(adminRef);
        if (!adminSnap.exists()) {
          await setDoc(adminRef, {
            email: admin.email,
            name: admin.name,
            role: 'ADMIN',
            sector: 'Almoxarifado',
            lastLogin: null
          });
        }
      });

      // Purge poli.almoxarifado@gmail.com from Firestore users collection if present
      (async () => {
        try {
          const poliRef = doc(db, 'users', 'poli.almoxarifado@gmail.com');
          const poliSnap = await getDoc(poliRef);
          if (poliSnap.exists()) {
            await deleteDoc(poliRef);
          }
        } catch (err) {
          console.warn("Could not remove poli.almoxarifado@gmail.com:", err);
        }
      })();

      const qUsers = query(collection(db, 'users'), orderBy('name', 'asc'));
      unsubscribeUsers = onSnapshot(qUsers, (snapshot) => {
        setUsersList(snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as UserProfile)));
      }, (error) => {
        handleFirestoreError(error, OperationType.LIST, 'users');
      });
    }
    return () => unsubscribeUsers();
  }, [user, userProfile]);

  const handleGoogleLogin = async () => {
    const provider = new GoogleAuthProvider();
    setLoginLoading(true);
    try {
      await signInWithPopup(auth, provider);
    } catch (error: any) {
      console.error("Login error:", error);
      if (error.code === 'auth/popup-blocked' || error.code === 'auth/popup-closed-by-user') {
        try {
          showToast("Redirecionando para login do Google...", "info");
          await signInWithRedirect(auth, provider);
          return;
        } catch (redirectErr: any) {
          console.error("Redirect login error:", redirectErr);
          showToast("O popup foi bloqueado pelo seu navegador. Por favor, permita janelas pop-up ou abra o sistema em uma nova aba.", "error");
        }
      } else if (error.code === 'auth/unauthorized-domain') {
        showToast("Erro: DomÃ­nio nÃ£o autorizado no Firebase Auth.", "error");
      } else {
        showToast(`Erro ao entrar: ${error.message}`, "error");
      }
    } finally {
      setLoginLoading(false);
    }
  };

  const handleEmailLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!authEmail || !authPassword) {
      alert("Preencha todos os campos.");
      return;
    }
    setLoginLoading(true);
    try {
      await signInWithEmailAndPassword(auth, authEmail, authPassword);
    } catch (error: any) {
      console.error("Login error:", error);
      alert(`Erro ao entrar: ${error.message}`);
    } finally {
      setLoginLoading(false);
    }
  };

  const handleRegister = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!authEmail || !authName || authSectors.length === 0) {
      showToast("Preencha todos os campos e selecione ao menos um setor.", "error");
      return;
    }
    setLoginLoading(true);
    try {
      // Just create/update the document in Firestore using email as ID
      // This allows the user to log in via Google later
      const userDocId = authEmail.toLowerCase().trim();
      const role = userDocId === 'gerlianemagalhaes79@gmail.com' ? 'ADMIN' : 'SETOR';
      
      await setDoc(doc(db, 'users', userDocId), {
        email: userDocId,
        name: authName,
        role: role,
        sector: authSectors[0], // Main sector or legacy
        allowedSectors: authSectors,
        registeredAt: new Date().toISOString()
      }, { merge: true });
      
      showToast("UsuÃ¡rio prÃ©-cadastrado com sucesso! Agora ele pode entrar usando o Google.", "success");
      setAuthEmail('');
      setAuthName('');
      setAuthSectors([]);
      setIsRegistering(false);
    } catch (error: any) {
      handleFirestoreError(error, OperationType.WRITE, 'users');
      showToast(`Erro ao cadastrar: ${error.message}`, "error");
    } finally {
      setLoginLoading(false);
    }
  };

  const handleLogout = () => signOut(auth);

  const handleDeleteTransaction = async (id: string, reason: string) => {
    if (!id) return;
    try {
      await runTransaction(db, async (transaction) => {
        const transRef = doc(db, 'transactions', id);
        const transSnap = await transaction.get(transRef);
        
        if (!transSnap.exists()) {
          throw new Error("MovimentaÃ§Ã£o nÃ£o encontrada.");
        }
        
        const transData = transSnap.data() as Transaction;

        if (transData.deletedAt) {
          throw new Error("Esta movimentaÃ§Ã£o jÃ¡ foi excluÃ­da.");
        }

        if (transData.item_id) {
          const itemRef = doc(db, 'items', transData.item_id);
          const itemSnap = await transaction.get(itemRef);
          
          if (itemSnap.exists()) {
            const itemData = itemSnap.data() as Item;
            const qty = Number(transData.quantity) || 0;
            let currentQty = Number(itemData.quantity) || 0;
            
            let newQty;
            if (transData.type === 'entry') {
              newQty = currentQty - qty;
            } else {
              newQty = currentQty + qty;
            }
            
            transaction.update(itemRef, { 
              quantity: Math.max(0, newQty),
              updatedAt: serverTimestamp()
            });
          }
        }

        transaction.update(transRef, {
          deletedAt: new Date().toISOString(),
          deletionReason: reason || 'Sem justificativa',
          deletedByEmail: user?.email
        });
      });

      setShowDeleteModal({ show: false });
      const itemName = transactions.find(t => t.id === id)?.item_name;
      setDeletionReason('');
      if (itemName) await checkStockAndNotify(itemName);
    } catch (error: any) {
      console.error("Error deleting transaction:", error);
      alert(`Erro ao excluir movimentaÃ§Ã£o: ${error.message}`);
    }
  };

  const handleRecoverTransaction = async (id: string) => {
    if (!id) return;
    try {
      await runTransaction(db, async (transaction) => {
        const transRef = doc(db, 'transactions', id);
        const transSnap = await transaction.get(transRef);
        
        if (!transSnap.exists()) {
          throw new Error("MovimentaÃ§Ã£o nÃ£o encontrada.");
        }
        
        const transData = transSnap.data() as Transaction;

        if (!transData.deletedAt) {
          throw new Error("Esta movimentaÃ§Ã£o nÃ£o estÃ¡ excluÃ­da.");
        }

        if (transData.item_id) {
          const itemRef = doc(db, 'items', transData.item_id);
          const itemSnap = await transaction.get(itemRef);
          
          if (itemSnap.exists()) {
            const itemData = itemSnap.data() as Item;
            const qty = Number(transData.quantity) || 0;
            let currentQty = Number(itemData.quantity) || 0;
            
            let newQty;
            if (transData.type === 'entry') {
              newQty = currentQty + qty;
            } else {
              newQty = currentQty - qty;
            }
            
            transaction.update(itemRef, { 
              quantity: Math.max(0, newQty),
              updatedAt: serverTimestamp()
            });
          }
        }

        transaction.update(transRef, {
          deletedAt: null,
          deletionReason: null,
          deletedByEmail: null
        });
      });
    } catch (error: any) {
      console.error("Error recovering transaction:", error);
      alert(`Erro ao recuperar movimentaÃ§Ã£o: ${error.message}`);
    }
  };

  const handleRecoverAllTransactions = async () => {
    const deletedTrans = transactions.filter(t => !!t.deletedAt);
    if (deletedTrans.length === 0) return;
    
    if (!confirm(`Deseja restaurar todas as ${deletedTrans.length} movimentaÃ§Ãµes excluÃ­das?`)) return;

    try {
      // We'll process them one by one to ensure stock is updated correctly via transactions
      for (const t of deletedTrans) {
        await handleRecoverTransaction(t.id);
      }
      alert("Todas as movimentaÃ§Ãµes foram restauradas com sucesso!");
    } catch (error: any) {
      console.error("Error recovering all transactions:", error);
      alert(`Erro ao restaurar movimentaÃ§Ãµes: ${error.message}`);
    }
  };

  const handleSubmitRequest = async () => {
    if (requestBasket.length === 0) {
      showToast("Adicione pelo menos um item Ã  solicitaÃ§Ã£o.", "error");
      return;
    }

    setIsSubmittingRequest(true);
    const loadingToast = showToast("Processando sua solicitaÃ§Ã£o...", "info");
    
    try {
      // 1. Fetch fresh inventory to validate stock correctly
      const itemsSnapshot = await getDocs(collection(db, 'items'));
      const freshItems = itemsSnapshot.docs
        .map(d => ({ id: d.id, ...d.data() } as Item))
        .filter(i => !i.deletedAt);
      
      // Calculate total stock with normalized names
      const totalInventory: Record<string, number> = {};
      freshItems.forEach(item => {
        if (!item.name) return;
        const key = normalizeString(item.name);
        totalInventory[key] = (totalInventory[key] || 0) + (Number(item.quantity) || 0);
      });

      // Aggregate current request basket quantities by product
      const basketAggregation: Record<string, number> = {};
      requestBasket.forEach(item => {
        const key = normalizeString(item.product_name);
        const qty = Math.max(1, Math.floor(Number(item.quantity) || 1));
        basketAggregation[key] = (basketAggregation[key] || 0) + qty;
      });

      // Validate stock
      for (const [productNameKey, requestedQty] of Object.entries(basketAggregation)) {
        const totalAvailable = totalInventory[productNameKey] || 0;
        
        if (requestedQty > totalAvailable) {
          const originalName = requestBasket.find(i => normalizeString(i.product_name) === productNameKey)?.product_name || "Produto";
          console.warn(`Stock check failed for ${productNameKey}: requested ${requestedQty}, available ${totalAvailable}`);
          showToast(
            `Estoque insuficiente para "${originalName}". DisponÃ­vel: ${totalAvailable}.`, 
            "error"
          );
          setIsSubmittingRequest(false);
          return;
        }
      }

      // 2. Prepare Request Data
      const requestId = editingRequest ? editingRequest.id : doc(collection(db, 'requests')).id;
      const batch = writeBatch(db);

      const requestData: any = {
        sector: selectedSector,
        date: editingRequest ? editingRequest.date : new Date().toISOString(),
        status: 'PENDENTE',
        observation: requestObservation || '',
        requesterEmail: user?.email || '',
        updatedAt: serverTimestamp(),
        isNewFlow: editingRequest ? (editingRequest.isNewFlow || false) : true
      };

      if (editingRequest) {
        batch.update(doc(db, 'requests', requestId), requestData);
        // Better to fetch directly here to be absolutely sure we have current items
        const oldItemsSnap = await getDocs(query(collection(db, 'request_items'), where('request_id', '==', requestId)));
        oldItemsSnap.docs.forEach(d => batch.delete(d.ref));
      } else {
        requestData.createdAt = serverTimestamp();
        requestData.requesterName = user?.displayName || user?.email || 'UsuÃ¡rio';
        batch.set(doc(db, 'requests', requestId), requestData);
      }

      // 3. Add current basket items
      requestBasket.forEach(item => {
        const itemRef = doc(collection(db, 'request_items'));
        batch.set(itemRef, {
          request_id: requestId,
          product_id: item.product_id,
          product_name: item.product_name,
          quantity_requested: Math.max(1, Math.floor(Number(item.quantity) || 1)),
          quantity_approved: Math.max(1, Math.floor(Number(item.quantity) || 1))
        });
      });

      // 4. Commit everything
      await batch.commit();

      if (!editingRequest) {
        // Notifications only for NEW requests
        try {
          const adminQuery = query(collection(db, 'users'), where('role', '==', 'ADMIN'));
          const almoxQuery = query(collection(db, 'users'), where('sector', '==', 'Almoxarifado'));
          
          const [adminSnap, almoxSnap] = await Promise.all([getDocs(adminQuery), getDocs(almoxQuery)]);
          const notified = new Set<string>();
          
          const notify = (snap: any) => {
            snap.forEach((d: any) => {
              if (!notified.has(d.id)) {
                createNotification(d.id, 'Nova SolicitaÃ§Ã£o', `Setor ${selectedSector} enviou uma nova solicitaÃ§Ã£o.`, requestId);
                notified.add(d.id);
              }
            });
          };
          notify(adminSnap);
          notify(almoxSnap);
        } catch (notifErr) {
          console.warn("Falha ao enviar notificaÃ§Ãµes:", notifErr);
        }
      }

      showToast(editingRequest ? "AlteraÃ§Ãµes salvas com sucesso!" : "SolicitaÃ§Ã£o enviada com sucesso!", "success");
      setRequestBasket([]);
      setRequestObservation('');
      setEditingRequest(null);
      setActiveTab('my-requests');
    } catch (error: any) {
      console.error("Erro crÃ­tico ao salvar:", error);
      showToast(`NÃ£o foi possÃ­vel salvar: ${error.message}. Verifique sua conexÃ£o e tente novamente.`, "error");
    } finally {
      setIsSubmittingRequest(false);
    }
  };

  const handleEditRequest = (request: MaterialRequest) => {
    setSelectedSector(request.sector);
    const items = allRequestItems.filter(ri => ri.request_id === request.id);
    setRequestBasket(items.map(i => ({
      product_id: i.product_id,
      product_name: i.product_name,
      quantity: i.quantity_requested
    })));
    setRequestObservation(request.observation || '');
    setEditingRequest(request);
    setActiveTab('new-request');
  };

  useEffect(() => {
    if (isAdmin) {
      cleanupOldDeletedData();
    }
  }, [isAdmin]);

  const cleanupOldDeletedData = async () => {
    const threeDaysAgo = subDays(new Date(), 3);
    
    try {
      // Cleanup items already loaded in state that are deleted > 3 days ago
      const deletedItems = items.filter(i => i.deletedAt && new Date(i.deletedAt) < threeDaysAgo);
      for (const item of deletedItems) {
        if (item.id) await deleteDoc(doc(db, 'items', item.id));
      }
      
      // Cleanup requests already loaded in state that are deleted > 3 days ago
      const deletedRequests = requests.filter(r => r.deletedAt && new Date(r.deletedAt) < threeDaysAgo);
      for (const req of deletedRequests) {
        if (req.id) await deleteDoc(doc(db, 'requests', req.id));
      }

      // Cleanup transactions already loaded in state that are deleted > 3 days ago
      const deletedTrans = transactions.filter(t => t.deletedAt && new Date(t.deletedAt) < threeDaysAgo);
      for (const t of deletedTrans) {
        if (t.id) await deleteDoc(doc(db, 'transactions', t.id));
      }
    } catch (error) {
      console.warn("Notice during cleanup:", error);
    }
  };

  const handleDeleteItem = async (itemId: string) => {
    if (!window.confirm("Tem certeza que deseja enviar este item para a lixeira? Ele serÃ¡ excluÃ­do definitivamente apÃ³s 3 dias.")) return;
    try {
      await updateDoc(doc(db, 'items', itemId), {
        deletedAt: new Date().toISOString(),
        deletedBy: user?.email
      });
      showToast("Item enviado para a lixeira.", "success");
    } catch (error: any) {
      handleFirestoreError(error, OperationType.UPDATE, `items/${itemId}`);
      showToast(`Erro ao excluir item: ${error.message}`, "error");
    }
  };

  const handlePrintRequests = async () => {
    const filteredRequests = requests.filter(req => {
      if (req.deletedAt || (req.status !== 'PENDENTE' && req.status !== 'EM_SEPARACAO')) return false;
      const reqDate = req.date.split('T')[0];
      return reqDate >= printRange.start && reqDate <= printRange.end;
    });

    if (filteredRequests.length === 0) {
      showToast("Nenhuma solicitaÃ§Ã£o pendente ou em separaÃ§Ã£o encontrada para este perÃ­odo.", "info");
      return;
    }

    try {
      const batch = writeBatch(db);
      let updatedAny = false;
      filteredRequests.forEach(req => {
        if (req.status === 'PENDENTE') {
          batch.update(doc(db, 'requests', req.id), {
            status: 'EM_SEPARACAO',
            updatedAt: serverTimestamp()
          });
          updatedAny = true;
        }
      });
      if (updatedAny) {
        await batch.commit();
        showToast("Status das solicitaÃ§Ãµes atualizado para 'Em SeparaÃ§Ã£o'!", "success");
      }
    } catch (error) {
      console.error("Erro ao atualizar status para EM_SEPARACAO:", error);
    }

    const printWindow = window.open('', '_blank');
    if (!printWindow) {
      showToast("Por favor, permita popups para imprimir.", "error");
      return;
    }

    const startDateStr = new Date(printRange.start + 'T12:00:00').toLocaleDateString('pt-BR');
    const endDateStr = new Date(printRange.end + 'T12:00:00').toLocaleDateString('pt-BR');
    const periodStr = printRange.start === printRange.end ? startDateStr : `${startDateStr} a ${endDateStr}`;

    const content = `
      <html>
        <head>
          <title>ImpressÃ£o de SolicitaÃ§Ãµes - ${periodStr}</title>
          <style>
            body { font-family: sans-serif; padding: 5px; color: #1C1917; font-size: 9px; line-height: 1.2; }
            .request-card { 
              border: 1px dashed #78716C; 
              border-radius: 6px; 
              padding: 8px; 
              margin-bottom: 12px; 
              page-break-inside: avoid;
            }
            .header-table { width: 100%; border-collapse: collapse; margin-bottom: 6px; }
            .header-table td { padding: 3px 5px; border: 1px solid #E7E5E4; font-size: 8.5px; }
            h1 { text-align: left; margin: 0 0 5px 0; font-size: 11px; text-transform: uppercase; border-bottom: 1.5px solid #1C1917; padding-bottom: 2px; }
            .items-table { width: 100%; border-collapse: collapse; margin-top: 6px; }
            .items-table th, .items-table td { border: 1px solid #1C1917; padding: 4px; text-align: left; font-size: 8.5px; }
            .items-table th { background-color: #FAFAF9; }
            .blank-col { width: 70px; text-align: center; }
            .footer { margin-top: 8px; text-align: center; font-size: 7px; color: #78716C; border-top: 1px dashed #E7E5E4; padding-top: 3px; }
            @media print {
              .no-print { display: none; }
            }
          </style>
        </head>
        <body>
          ${filteredRequests.map((req, idx) => {
            const items = allRequestItems.filter(ri => ri.request_id === req.id);
            return `
              <div class="request-card">
                <h1>SolicitaÃ§Ã£o de Material</h1>
                <table class="header-table">
                  <tr>
                    <td><strong>NÃºmero:</strong> #${req.id.slice(-5).toUpperCase()}</td>
                    <td><strong>Data de CriaÃ§Ã£o:</strong> ${new Date(req.date).toLocaleDateString('pt-BR')}</td>
                  </tr>
                  <tr>
                    <td><strong>Setor Solicitante:</strong> ${req.sector}</td>
                    <td><strong>Status:</strong> EM SEPARAÃ‡ÃƒO</td>
                  </tr>
                  <tr>
                    <td colspan="2"><strong>Solicitante:</strong> ${req.requesterEmail}</td>
                  </tr>
                  ${req.observation ? `<tr><td colspan="2"><strong>ObservaÃ§Ãµes:</strong> ${req.observation}</td></tr>` : ''}
                </table>

                <h3 style="margin: 6px 0 3px 0; font-size: 9px; border-bottom: 1px solid #1C1917; padding-bottom: 2px; text-transform: uppercase;">ITENS DA SOLICITAÃ‡ÃƒO (Para separaÃ§Ã£o fÃ­sica)</h3>
                <table class="items-table">
                  <thead>
                    <tr>
                      <th>Produto / DescriÃ§Ã£o</th>
                      <th style="width: 70px; text-align: center;">Qtd Solicitada</th>
                      <th class="blank-col">Qtd Separada</th>
                      <th>Obs. / Lote do Material</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${items.map(item => `
                      <tr>
                        <td style="font-weight: bold; font-size: 8.5px;">${item.product_name}</td>
                        <td style="text-align: center; font-size: 8.5px; font-weight: bold;">${item.quantity_requested}</td>
                        <td class="blank-col" style="border-bottom: 1px solid #1C1917;"></td>
                        <td style="border-bottom: 1px solid #1C1917;"></td>
                      </tr>
                    `).join('')}
                  </tbody>
                </table>

                <div class="footer">Gerado em ${new Date().toLocaleString('pt-BR')}</div>
              </div>
            `;
          }).join('')}
          <script>
            window.onload = () => {
              window.print();
            };
          </script>
        </body>
      </html>
    `;

    printWindow.document.write(content);
    printWindow.document.close();
  };

  const handleDeleteRequest = async (requestId: string) => {
    const reqToDel = requests.find(r => r.id === requestId);
    if (reqToDel?.status === 'ENTREGUE') {
      showToast("NÃ£o Ã© possÃ­vel excluir uma solicitaÃ§Ã£o que jÃ¡ foi entregue.", "error");
      return;
    }
    
    if (!window.confirm("Tem certeza que deseja enviar esta solicitaÃ§Ã£o para a lixeira? Ela serÃ¡ excluÃ­da definitivamente apÃ³s 3 dias.")) return;
    try {
      await updateDoc(doc(db, 'requests', requestId), {
        deletedAt: new Date().toISOString(),
        deletedBy: user?.email
      });
      showToast("SolicitaÃ§Ã£o enviada para a lixeira.", "success");
    } catch (error: any) {
      handleFirestoreError(error, OperationType.UPDATE, `requests/${requestId}`);
      showToast(`Erro ao excluir solicitaÃ§Ã£o: ${error.message}`, "error");
    }
  };

  const handleAddExtraItemToRequest = async (requestId: string, productName: string, productId: string) => {
    setIsAdminAddingItem(true);
    try {
      const newItem: Omit<RequestItem, 'id'> = {
        request_id: requestId,
        product_id: productId,
        product_name: productName,
        quantity_requested: 1,
        quantity_approved: 1
      };
      
      await addDoc(collection(db, 'request_items'), newItem);
      setAdminAddItemSearch('');
      showToast(`"${productName}" adicionado Ã  solicitaÃ§Ã£o.`, "success");
    } catch (error: any) {
      handleFirestoreError(error, OperationType.CREATE, 'request_items');
      showToast("Erro ao adicionar item.", "error");
    } finally {
      setIsAdminAddingItem(false);
    }
  };

  const handleUpdateObservation = async (requestId: string) => {
    try {
      await updateDoc(doc(db, 'requests', requestId), { 
        adminObservation: adminObservation 
      });
      showToast("ObservaÃ§Ã£o atualizada com sucesso!", "success");
    } catch (error: any) {
      console.error("Error updating observation:", error);
      showToast(`Erro ao atualizar observaÃ§Ã£o: ${error.message}`, "error");
    }
  };

  const handlePrintSingleRequest = async (request: MaterialRequest) => {
    // 1. If the request is in PENDENTE state, transition it to EM_SEPARACAO
    if (request.isNewFlow && request.status === 'PENDENTE') {
      try {
        await updateDoc(doc(db, 'requests', request.id), {
          status: 'EM_SEPARACAO',
          updatedAt: serverTimestamp()
        });
        showToast("Status alterado para 'Em SeparaÃ§Ã£o'!", "success");
        // Update local modal state immediately
        if (showRequestDetailModal.show && showRequestDetailModal.request?.id === request.id) {
          setShowRequestDetailModal({
            ...showRequestDetailModal,
            request: { ...showRequestDetailModal.request, status: 'EM_SEPARACAO' }
          });
        }
      } catch (error) {
        console.error("Erro ao atualizar status para EM_SEPARACAO:", error);
      }
    }

    const printWindow = window.open('', '_blank');
    if (!printWindow) {
      showToast("Por favor, permita popups para imprimir.", "error");
      return;
    }

    const items = allRequestItems.filter(ri => ri.request_id === request.id);
    const dateStr = new Date(request.date).toLocaleDateString('pt-BR');

    const content = `
      <html>
        <head>
          <title>SolicitaÃ§Ã£o de Material - #${request.id.slice(-5).toUpperCase()}</title>
          <style>
            body { font-family: sans-serif; padding: 20px; color: #1C1917; }
            .header-table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
            .header-table td { padding: 8px; border: 1px solid #E7E5E4; }
            h1 { text-align: center; margin-bottom: 20px; font-size: 22px; text-transform: uppercase; border-bottom: 3px double #1C1917; padding-bottom: 10px; }
            .items-table { width: 100%; border-collapse: collapse; margin-top: 20px; }
            .items-table th, .items-table td { border: 1px solid #1C1917; padding: 10px; text-align: left; font-size: 13px; }
            .items-table th { background-color: #FAFAF9; }
            .blank-col { width: 120px; text-align: center; }
            .signature-section { margin-top: 60px; display: flex; justify-content: space-between; }
            .signature-box { width: 45%; text-align: center; border-top: 1px solid #1C1917; padding-top: 5px; font-size: 12px; }
            .footer { margin-top: 50px; text-align: center; font-size: 10px; color: #78716C; border-top: 1px solid #E7E5E4; padding-top: 10px; }
            @media print {
              .no-print { display: none; }
            }
          </style>
        </head>
        <body>
          <h1>SolicitaÃ§Ã£o de Material</h1>
          <table class="header-table">
            <tr>
              <td><strong>NÃºmero:</strong> #${request.id.slice(-5).toUpperCase()}</td>
              <td><strong>Data:</strong> ${dateStr}</td>
            </tr>
            <tr>
              <td><strong>Setor Solicitante:</strong> ${request.sector}</td>
              <td><strong>Status:</strong> ${request.status === 'PENDENTE' ? 'PENDENTE' : 'EM SEPARAÃ‡ÃƒO'}</td>
            </tr>
            <tr>
              <td colspan="2"><strong>Solicitante:</strong> ${request.requesterEmail}</td>
            </tr>
            ${request.observation ? `<tr><td colspan="2"><strong>ObservaÃ§Ãµes do Solicitante:</strong> ${request.observation}</td></tr>` : ''}
          </table>

          <h3 style="margin-top: 30px; font-size: 16px; border-bottom: 1px solid #1C1917; padding-bottom: 5px;">ITENS DA SOLICITAÃ‡ÃƒO (Para separaÃ§Ã£o fÃ­sica)</h3>
          <table class="items-table">
            <thead>
              <tr>
                <th>Produto / DescriÃ§Ã£o</th>
                <th style="width: 100px; text-align: center;">Qtd Solicitada</th>
                <th class="blank-col">Qtd Separada (Anotar)</th>
                <th>Obs. / Lote do Material</th>
              </tr>
            </thead>
            <tbody>
              ${items.map(item => `
                <tr>
                  <td style="font-weight: bold;">${item.product_name}</td>
                  <td style="text-align: center; font-size: 14px; font-weight: bold;">${item.quantity_requested}</td>
                  <td class="blank-col" style="border-bottom: 1px solid #1C1917;"></td>
                  <td style="border-bottom: 1px solid #1C1917;"></td>
                </tr>
              `).join('')}
            </tbody>
          </table>

          <div class="signature-section">
            <div class="signature-box" style="margin-top: 40px;">
              <br/><br/>
              ________________________________________<br/>
              Setor Solicitante (Assinatura de Recebimento)
            </div>
            <div class="signature-box" style="margin-top: 40px;">
              <br/><br/>
              ________________________________________<br/>
              ResponsÃ¡vel pela SeparaÃ§Ã£o (Almoxarifado)
            </div>
          </div>

          <div class="footer">Gerado via Sistema de Almoxarifado em ${new Date().toLocaleString('pt-BR')}</div>
          <script>
            window.onload = () => {
              window.print();
            };
          </script>
        </body>
      </html>
    `;

    printWindow.document.write(content);
    printWindow.document.close();
  };

  const handleApproveAndDeliverNewRequest = async (requestId: string, currentRequestItems: RequestItem[]) => {
    try {
      showToast("Processando aprovaÃ§Ã£o e baixa no estoque...", "info");
      
      const requestRef = doc(db, 'requests', requestId);
      const requestSnap = await getDoc(requestRef);
      if (!requestSnap.exists()) throw new Error("SolicitaÃ§Ã£o nÃ£o encontrada.");
      const requestData = requestSnap.data() as MaterialRequest;

      if (requestData.status === 'ENTREGUE') {
        showToast("Esta solicitaÃ§Ã£o jÃ¡ foi entregue.", "info");
        return;
      }

      // Pre-fetch all necessary stock data with normalized name matching
      const itemsSnapshot = await getDocs(collection(db, 'items'));
      const allActiveItems = itemsSnapshot.docs
        .map(d => ({ id: d.id, ...d.data() } as Item))
        .filter(i => !i.deletedAt);

      const itemsStockData: any[] = [];
      for (const reqItem of currentRequestItems) {
        if (reqItem.quantity_approved <= 0) continue;

        const normalizedReqName = normalizeString(reqItem.product_name);
        
        // Find all batches that represent this product (same normalized name)
        let batches = allActiveItems.filter(item => 
          normalizeString(item.name) === normalizedReqName && (item.quantity || 0) > 0
        );

        batches.sort((a, b) => {
          if (a.expiry_date === 'Indeterminada' || !a.expiry_date) return 1;
          if (b.expiry_date === 'Indeterminada' || !b.expiry_date) return -1;
          return new Date(a.expiry_date).getTime() - new Date(b.expiry_date).getTime();
        });

        let pharmItems: any[] = [];
        if (requestData.sector === 'FarmÃ¡cia') {
          pharmItems = allActiveItems
            .filter(item => normalizeString(item.name) === normalizedReqName && item.location === 'FarmÃ¡cia')
            .map(item => ({ id: item.id, batch_number: item.batch_number, ref: doc(db, 'items', item.id) }));
        }

        itemsStockData.push({ reqItem, batches, pharmItems });
      }

      await runTransaction(db, async (transaction) => {
        // Collect all batch and pharmacy refs to read them all first
        const batchRefs = itemsStockData.flatMap(d => d.batches.map(b => doc(db, 'items', b.id)));
        const pharmRefs = itemsStockData.flatMap(d => d.pharmItems.map(p => p.ref));
        
        // 1. Perform ALL reads first
        const [tRequestSnap, ...itemSnaps] = await Promise.all([
          transaction.get(requestRef),
          ...batchRefs.map(ref => transaction.get(ref)),
          ...pharmRefs.map(ref => transaction.get(ref))
        ]);

        const tRequestData = tRequestSnap.data() as MaterialRequest | undefined;
        if (!tRequestData || tRequestData.status === 'ENTREGUE') return;

        // Map snapshots for easy access by path
        const snapMap = new Map();
        itemSnaps.forEach(snap => snapMap.set(snap.ref.path, snap));

        // 2. Perform ALL writes
        // First update the main request document
        transaction.update(requestRef, { 
          status: 'ENTREGUE',
          adminObservation: adminObservation,
          deliveredAt: new Date().toISOString(),
          deliveredBy: user?.email,
          updatedAt: serverTimestamp()
        });

        // Also update all the request_items quantity_approved in the database
        currentRequestItems.forEach(item => {
          const itemRef = doc(db, 'request_items', item.id);
          transaction.update(itemRef, { quantity_approved: item.quantity_approved });
        });

        for (const { reqItem, batches, pharmItems } of itemsStockData) {
          let remaining = reqItem.quantity_approved;
          
          for (const batch of batches) {
            if (remaining <= 0) break;

            const tBatchRef = doc(db, 'items', batch.id);
            const tBatchSnap = snapMap.get(tBatchRef.path);
            if (!tBatchSnap || !tBatchSnap.exists()) continue;
            
            const tBatchData = tBatchSnap.data() as Item;
            const currentQty = tBatchData.quantity || 0;
            if (currentQty <= 0) continue;

            const toTake = Math.min(currentQty, remaining);
            
            transaction.update(tBatchRef, {
              quantity: currentQty - toTake,
              updatedAt: serverTimestamp()
            });

            // Log Transaction
            const transRef = doc(collection(db, 'transactions'));
            transaction.set(transRef, {
              item_id: batch.id,
              item_name: reqItem.product_name,
              type: 'exit',
              origin: batch.origin || 'extra',
              quantity: toTake,
              sector: requestData.sector,
              location: batch.location || 'Almoxarifado',
              date: new Date().toISOString(),
              responsible: user?.displayName || user?.email,
              responsibleEmail: user?.email,
              exitReason: 'consumo',
              batch_number: batch.batch_number,
              expiry_date: batch.expiry_date
            });

            if (requestData.sector === 'FarmÃ¡cia' && batch.location !== 'FarmÃ¡cia') {
              const existingPharm = pharmItems.find((p: any) => p.batch_number === batch.batch_number);
              if (existingPharm) {
                const tPharmRef = existingPharm.ref;
                const tPharmSnap = snapMap.get(tPharmRef.path);
                const tPharmData = tPharmSnap?.data() as Item | undefined;
                transaction.update(tPharmRef, {
                  quantity: (tPharmData?.quantity || 0) + toTake,
                  updatedAt: serverTimestamp()
                });
              } else {
                const newItemRef = doc(collection(db, 'items'));
                transaction.set(newItemRef, {
                  name: reqItem.product_name,
                  description: batch.description || '',
                  category: batch.category || 'Outros',
                  supplier: batch.supplier || 'TransferÃªncia',
                  batch_number: batch.batch_number || '',
                  expiry_date: batch.expiry_date || 'Indeterminada',
                  initial_quantity: toTake,
                  quantity: toTake,
                  min_quantity: batch.min_quantity || 0,
                  unit_price: batch.unit_price || 0,
                  location: 'FarmÃ¡cia',
                  origin: batch.origin || 'extra',
                  date: new Date().toISOString(),
                  createdAt: serverTimestamp(),
                  updatedAt: serverTimestamp()
                });
              }
            }
            remaining -= toTake;
          }

          if (remaining > 0) {
            throw new Error(`Estoque insuficiente para "${reqItem.product_name}".`);
          }
        }
      });

      // Cleanup and UI updates
      showToast("SolicitaÃ§Ã£o aprovada, entregue e estoque baixado automaticamente!", "success");
      setShowRequestDetailModal({ show: false });

      // Notifications
      const uSnap = await getDocs(query(collection(db, 'users'), where('email', '==', requestData.requesterEmail)));
      if (!uSnap.empty) {
        await createNotification(uSnap.docs[0].id, 'SolicitaÃ§Ã£o Entregue', `Sua solicitaÃ§Ã£o #${requestId.slice(-5).toUpperCase()} foi aprovada e entregue.`, requestId);
      }

      // Stock Zero Notifications
      for (const { reqItem } of itemsStockData) {
        await checkStockAndNotify(reqItem.product_name);
      }

      // Receipt
      const itemsForReceipt: { product_name: string; quantity: number; batch_number: string }[] = [];
      for (const { reqItem, batches } of itemsStockData) {
        let remaining = reqItem.quantity_approved;
        for (const batch of batches) {
          if (remaining <= 0) break;
          const available = batch.quantity || 0;
          if (available <= 0) continue;
          const toTake = Math.min(available, remaining);
          itemsForReceipt.push({
            product_name: reqItem.product_name,
            quantity: toTake,
            batch_number: batch.batch_number || '---'
          });
          remaining -= toTake;
        }
        if (remaining > 0 && !itemsForReceipt.some(i => i.product_name === reqItem.product_name)) {
          itemsForReceipt.push({
            product_name: reqItem.product_name,
            quantity: reqItem.quantity_approved,
            batch_number: '---'
          });
        }
      }
      if (itemsForReceipt.length > 0) {
        handleExportDeliveryReceiptPDF({
          sector: requestData.sector,
          items: itemsForReceipt,
          requestId: requestId,
          date: new Date().toISOString()
        });
      }

    } catch (error: any) {
      console.error("Erro ao aprovar e entregar:", error);
      showToast(`Erro no processo: ${error.message}`, "error");
    }
  };

  const handleRequestDevolution = async () => {
    if (devolutionBasket.length === 0) {
      showToast("Por favor, adicione pelo menos um item Ã  devoluÃ§Ã£o.", "info");
      return;
    }

    // Validate quantities
    for (const item of devolutionBasket) {
      if (item.quantity <= 0) {
        showToast(`Por favor, insira uma quantidade maior que zero para ${item.product_name}.`, "error");
        return;
      }
      if (item.quantity > item.maxQty) {
        showToast(`Quantidade invÃ¡lida para ${item.product_name}. MÃ¡ximo permitido: ${item.maxQty}`, "error");
        return;
      }
    }

    try {
      setIsProcessingDevolution(true);
      showToast("Enviando solicitaÃ§Ã£o de devoluÃ§Ã£o...", "info");

      const batch = writeBatch(db);
      const newReqRef = doc(collection(db, 'requests'));
      
      const requestData = {
        sector: selectedSector,
        date: new Date().toISOString(),
        status: 'DEVOLUCAO_PENDENTE',
        isReturn: true,
        originalRequestId: showDevolutionModal.request?.id || '',
        returnReason: devolutionReason,
        observation: devolutionObservation || '',
        requesterEmail: user?.email || '',
        requesterName: userProfile?.name || user?.displayName || user?.email || 'UsuÃ¡rio',
        isNewFlow: true,
        createdAt: serverTimestamp(),
        updatedAt: serverTimestamp()
      };

      batch.set(newReqRef, requestData);

      devolutionBasket.forEach((item) => {
        const itemRef = doc(collection(db, 'request_items'));
        const productBatches = items.filter(i => !i.deletedAt && i.name.trim().toLowerCase() === item.product_name.trim().toLowerCase());
        const validBatchId = (item.selectedBatchId && productBatches.some(b => b.id === item.selectedBatchId))
          ? item.selectedBatchId
          : (productBatches[0]?.id || '');

        batch.set(itemRef, {
          request_id: newReqRef.id,
          product_id: item.product_id,
          product_name: item.product_name,
          quantity_requested: item.quantity,
          quantity_approved: item.quantity,
          batch_id: validBatchId
        });
      });

      // Notify administrators
      try {
        const adminSnap = await getDocs(query(collection(db, 'users'), where('role', '==', 'ADMIN')));
        adminSnap.forEach(adminDoc => {
          const notifRef = doc(collection(db, 'notifications'));
          batch.set(notifRef, {
            userId: adminDoc.id,
            title: 'SolicitaÃ§Ã£o de DevoluÃ§Ã£o',
            message: `Setor ${selectedSector} solicitou devoluÃ§Ã£o de materiais.`,
            date: new Date().toISOString(),
            read: false,
            requestId: newReqRef.id,
            type: 'REQUEST'
          });
        });
      } catch (e) {
        console.warn("Aviso ao notificar administradores:", e);
      }

      await batch.commit();

      showToast("SolicitaÃ§Ã£o de devoluÃ§Ã£o enviada para o almoxarifado!", "success");
      setShowDevolutionModal({ show: false });
      setDevolutionBasket([]);
      setDevolutionObservation('');
      
      if (showRequestDetailModal.show && showDevolutionModal.request && showRequestDetailModal.request?.id === showDevolutionModal.request.id) {
        setShowRequestDetailModal({ show: false });
      }

    } catch (error: any) {
      console.error("Erro ao solicitar devoluÃ§Ã£o:", error);
      showToast(`Erro ao solicitar devoluÃ§Ã£o: ${error.message}`, "error");
    } finally {
      setIsProcessingDevolution(false);
    }
  };

  const handleApproveDevolution = async (requestId: string, devItems: RequestItem[]) => {
    try {
      setIsProcessingDevolution(true);
      showToast("Aprovando devoluÃ§Ã£o e retornando ao estoque...", "info");

      // Fetch active stock items
      const itemsSnapshot = await getDocs(collection(db, 'items'));
      const allActiveItems = itemsSnapshot.docs
        .map(d => ({ id: d.id, ...d.data() } as Item))
        .filter(i => !i.deletedAt);

      await runTransaction(db, async (transaction) => {
        const requestRef = doc(db, 'requests', requestId);
        const requestSnap = await transaction.get(requestRef);
        if (!requestSnap.exists()) throw new Error("SolicitaÃ§Ã£o de devoluÃ§Ã£o nÃ£o encontrada.");
        const requestData = requestSnap.data() as MaterialRequest;

        if (requestData.status === 'DEVOLUCAO_APROVADA') {
          throw new Error("Esta devoluÃ§Ã£o jÃ¡ foi aprovada anteriormente.");
        }

        // Collect all doc IDs that we need to read in the transaction:
        // 1) Sector source items (FarmÃ¡cia / Requesting sector)
        // 2) Almoxarifado target items
        const docIdsToRead = new Set<string>();

        for (const item of devItems) {
          const returnQty = item.quantity_approved || item.quantity_requested || 0;
          if (returnQty <= 0) continue;

          // Find source item in sector stock
          if (item.batch_id && allActiveItems.some(i => i.id === item.batch_id)) {
            docIdsToRead.add(item.batch_id);
          }
          const sectorItem = allActiveItems.find(i => 
            i.name.trim().toLowerCase() === item.product_name.trim().toLowerCase() && 
            i.location === requestData.sector
          );
          if (sectorItem) {
            docIdsToRead.add(sectorItem.id);
          }

          // Find target item in Almoxarifado stock
          const almoxItem = allActiveItems.find(i => 
            i.name.trim().toLowerCase() === item.product_name.trim().toLowerCase() && 
            (!i.location || i.location === 'Almoxarifado')
          );
          if (almoxItem) {
            docIdsToRead.add(almoxItem.id);
          }
        }

        // Read all docs inside transaction
        const snapMap = new Map<string, any>();
        for (const id of docIdsToRead) {
          const itemRef = doc(db, 'items', id);
          const snap = await transaction.get(itemRef);
          snapMap.set(id, snap);
        }

        // Now perform transaction writes
        for (const item of devItems) {
          const returnQty = item.quantity_approved || item.quantity_requested || 0;
          if (returnQty <= 0) continue;

          // 1. DECREASE stock in sector (e.g., FarmÃ¡cia)
          let sourceItemDoc: { id: string, data: Item } | undefined;

          // Check if item.batch_id is a valid sector item
          if (item.batch_id && snapMap.has(item.batch_id)) {
            const snap = snapMap.get(item.batch_id);
            if (snap && snap.exists()) {
              const data = snap.data() as Item;
              if (data.location === requestData.sector) {
                sourceItemDoc = { id: item.batch_id, data };
              }
            }
          }

          // Fallback search for sector item by name & location
          if (!sourceItemDoc) {
            for (const [id, snap] of snapMap.entries()) {
              if (snap && snap.exists()) {
                const data = snap.data() as Item;
                if (data.location === requestData.sector && data.name.trim().toLowerCase() === item.product_name.trim().toLowerCase()) {
                  sourceItemDoc = { id, data };
                  break;
                }
              }
            }
          }

          if (sourceItemDoc) {
            const sourceRef = doc(db, 'items', sourceItemDoc.id);
            const currentQty = Number(sourceItemDoc.data.quantity) || 0;
            const newQty = Math.max(0, currentQty - returnQty);
            transaction.update(sourceRef, {
              quantity: newQty,
              updatedAt: serverTimestamp()
            });
            sourceItemDoc.data.quantity = newQty; // update in-memory
          }

          // 2. INCREASE stock in Almoxarifado
          let almoxItemDoc: { id: string, data: Item } | undefined;
          for (const [id, snap] of snapMap.entries()) {
            if (snap && snap.exists()) {
              const data = snap.data() as Item;
              if ((!data.location || data.location === 'Almoxarifado') && data.name.trim().toLowerCase() === item.product_name.trim().toLowerCase()) {
                almoxItemDoc = { id, data };
                break;
              }
            }
          }

          let almoxRef: DocumentReference;
          let batchNumber = sourceItemDoc?.data.batch_number || 'DevoluÃ§Ã£o';
          let expiryDate = sourceItemDoc?.data.expiry_date || 'Indeterminada';
          let category = sourceItemDoc?.data.category || 'Medicamentos';
          let unitMeasure = sourceItemDoc?.data.unit_measure || 'Unidade (UN)';

          if (almoxItemDoc) {
            almoxRef = doc(db, 'items', almoxItemDoc.id);
            const currentAlmoxQty = Number(almoxItemDoc.data.quantity) || 0;
            const newAlmoxQty = currentAlmoxQty + returnQty;
            transaction.update(almoxRef, {
              quantity: newAlmoxQty,
              updatedAt: serverTimestamp()
            });
            almoxItemDoc.data.quantity = newAlmoxQty; // update in-memory
          } else {
            // Create new stock item in Almoxarifado if none exists
            const newStockRef = doc(collection(db, 'items'));
            almoxRef = newStockRef;
            transaction.set(newStockRef, {
              name: item.product_name,
              quantity: returnQty,
              min_quantity: 10,
              category: category,
              unit: 'unid',
              unit_measure: unitMeasure,
              location: 'Almoxarifado',
              origin: 'extra',
              batch_number: batchNumber,
              expiry_date: expiryDate,
              entry_date: new Date().toISOString(),
              createdAt: serverTimestamp(),
              updatedAt: serverTimestamp()
            });
          }

          // 3. Log entry transaction for Almoxarifado
          const transRefEntry = doc(collection(db, 'transactions'));
          transaction.set(transRefEntry, {
            item_id: almoxRef.id,
            item_name: item.product_name,
            type: 'entry',
            origin: 'extra',
            quantity: returnQty,
            sector: requestData.sector,
            location: 'Almoxarifado',
            date: new Date().toISOString(),
            responsible: userProfile?.name || user?.displayName || user?.email || 'Administrador',
            responsibleEmail: user?.email,
            batch_number: batchNumber,
            expiry_date: expiryDate,
            isReturn: true,
            returnReason: requestData.returnReason || 'NÃ£o especificado',
            observation: requestData.observation || ''
          });

          // 4. Log exit transaction for Sector/FarmÃ¡cia
          const transRefExit = doc(collection(db, 'transactions'));
          transaction.set(transRefExit, {
            item_id: sourceItemDoc ? sourceItemDoc.id : almoxRef.id,
            item_name: item.product_name,
            type: 'exit',
            origin: 'extra',
            quantity: returnQty,
            sector: requestData.sector,
            location: requestData.sector,
            date: new Date().toISOString(),
            responsible: userProfile?.name || user?.displayName || user?.email || 'Administrador',
            responsibleEmail: user?.email,
            exitReason: 'vencido',
            expiryReason: requestData.returnReason || 'DevoluÃ§Ã£o ao Almoxarifado',
            batch_number: batchNumber,
            expiry_date: expiryDate,
            isReturn: true
          });
        }

        // Update main request status
        transaction.update(requestRef, {
          status: 'DEVOLUCAO_APROVADA',
          adminObservation: adminObservation || '',
          approvedAt: new Date().toISOString(),
          approvedBy: user?.email || 'Administrador',
          updatedAt: serverTimestamp()
        });
      });

      // Post-transaction updates: update quantity_returned on original request items
      const requestSnap = await getDoc(doc(db, 'requests', requestId));
      const requestData = requestSnap?.data() as MaterialRequest | undefined;
      if (requestData) {
        const batch = writeBatch(db);
        if (requestData.originalRequestId) {
          const origItemsSnap = await getDocs(query(collection(db, 'request_items'), where('request_id', '==', requestData.originalRequestId)));
          origItemsSnap.docs.forEach(d => {
            const origItem = d.data() as RequestItem;
            const matchedDevItem = devItems.find(di => di.product_name.trim().toLowerCase() === origItem.product_name.trim().toLowerCase());
            if (matchedDevItem) {
              const returnQty = matchedDevItem.quantity_approved || matchedDevItem.quantity_requested || 0;
              const currentReturned = origItem.quantity_returned || 0;
              batch.update(d.ref, {
                quantity_returned: currentReturned + returnQty
              });
            }
          });
        } else {
          // Direct flow: find and update ENTREGUE requests for this sector matching product_name
          const sectorReqsSnap = await getDocs(query(
            collection(db, 'requests'), 
            where('sector', '==', requestData.sector),
            where('status', '==', 'ENTREGUE')
          ));
          const sectorReqIds = sectorReqsSnap.docs.map(d => d.id);
          if (sectorReqIds.length > 0) {
            for (const matchedDevItem of devItems) {
              const returnQty = matchedDevItem.quantity_approved || matchedDevItem.quantity_requested || 0;
              if (returnQty <= 0) continue;

              const origItemsSnap = await getDocs(query(
                collection(db, 'request_items'),
                where('product_name', '==', matchedDevItem.product_name)
              ));

              let remainingToDistribute = returnQty;
              for (const d of origItemsSnap.docs) {
                const origItem = d.data() as RequestItem;
                if (sectorReqIds.includes(origItem.request_id)) {
                  const maxCanReturn = origItem.quantity_approved - (origItem.quantity_returned || 0);
                  if (maxCanReturn > 0 && remainingToDistribute > 0) {
                    const toReturn = Math.min(maxCanReturn, remainingToDistribute);
                    batch.update(d.ref, {
                      quantity_returned: (origItem.quantity_returned || 0) + toReturn
                    });
                    remainingToDistribute -= toReturn;
                  }
                }
              }
            }
          }
        }
        await batch.commit();
      }

      // Notify the requester
      if (requestData && requestData.requesterEmail) {
        const userSnap = await getDocs(query(collection(db, 'users'), where('email', '==', requestData.requesterEmail)));
        if (!userSnap.empty) {
          await addDoc(collection(db, 'notifications'), {
            userId: userSnap.docs[0].id,
            title: 'DevoluÃ§Ã£o Aprovada',
            message: `Sua solicitaÃ§Ã£o de devoluÃ§Ã£o para o setor ${requestData.sector} foi aprovada. Os materiais retornaram ao estoque.`,
            date: new Date().toISOString(),
            read: false,
            requestId: requestId,
            type: 'REQUEST'
          });
        }
      }

      showToast("DevoluÃ§Ã£o aprovada com sucesso! Materiais retornados ao estoque.", "success");
      setShowRequestDetailModal({ show: false });

    } catch (error: any) {
      console.error("Erro ao aprovar devoluÃ§Ã£o:", error);
      showToast(`Erro ao aprovar devoluÃ§Ã£o: ${error.message}`, "error");
    } finally {
      setIsProcessingDevolution(false);
    }
  };

  const handleRejectDevolution = async (requestId: string) => {
    try {
      setIsProcessingDevolution(true);
      showToast("Recusando devoluÃ§Ã£o...", "info");

      await updateDoc(doc(db, 'requests', requestId), {
        status: 'DEVOLUCAO_RECUSADA',
        adminObservation: adminObservation || '',
        updatedAt: serverTimestamp()
      });

      // Notify requester
      const requestSnap = await getDoc(doc(db, 'requests', requestId));
      const requestData = requestSnap.data() as MaterialRequest | undefined;
      if (requestData && requestData.requesterEmail) {
        const userSnap = await getDocs(query(collection(db, 'users'), where('email', '==', requestData.requesterEmail)));
        if (!userSnap.empty) {
          await addDoc(collection(db, 'notifications'), {
            userId: userSnap.docs[0].id,
            title: 'DevoluÃ§Ã£o Recusada',
            message: `Sua solicitaÃ§Ã£o de devoluÃ§Ã£o para o setor ${requestData.sector} foi recusada pelo almoxarifado.`,
            date: new Date().toISOString(),
            read: false,
            requestId: requestId,
            type: 'REQUEST'
          });
        }
      }

      showToast("DevoluÃ§Ã£o recusada com sucesso.", "success");
      setShowRequestDetailModal({ show: false });
    } catch (error: any) {
      console.error("Erro ao recusar devoluÃ§Ã£o:", error);
      showToast(`Erro ao recusar devoluÃ§Ã£o: ${error.message}`, "error");
    } finally {
      setIsProcessingDevolution(false);
    }
  };

  const handleApproveRequest = async (requestId: string, items: RequestItem[]) => {
    try {
      const batch = writeBatch(db);
      const requestRef = doc(db, 'requests', requestId);
      batch.update(requestRef, { 
        status: 'APROVADO',
        adminObservation: adminObservation,
        updatedAt: serverTimestamp()
      });
      
      items.forEach(item => {
        const itemRef = doc(db, 'request_items', item.id);
        batch.update(itemRef, { quantity_approved: item.quantity_approved });
      });

      await batch.commit();
      
      const request = requests.find(r => r.id === requestId);
      if (request) {
        const userSnap = await getDocs(query(collection(db, 'users'), where('email', '==', request.requesterEmail)));
        if (!userSnap.empty) {
          const msg = adminObservation 
            ? `Sua solicitaÃ§Ã£o #${requestId.slice(-5).toUpperCase()} foi aprovada. Obs: ${adminObservation}`
            : `Sua solicitaÃ§Ã£o #${requestId.slice(-5).toUpperCase()} foi aprovada.`;
          await createNotification(userSnap.docs[0].id, 'SolicitaÃ§Ã£o Aprovada', msg, requestId);
        }
      }

      showToast("SolicitaÃ§Ã£o aprovada!", "success");
      setShowRequestDetailModal({ show: false });
    } catch (error: any) {
      handleFirestoreError(error, OperationType.UPDATE, `requests/${requestId}`);
      showToast(`Erro ao aprovar: ${error.message}`, "error");
    }
  };

  const handleDeliverRequest = async (requestId: string, requestItems: RequestItem[]) => {
    try {
      showToast("Processando entrega... Aguarde.", "info");
      
      const requestRef = doc(db, 'requests', requestId);
      const requestSnap = await getDoc(requestRef);
      if (!requestSnap.exists()) throw new Error("SolicitaÃ§Ã£o nÃ£o encontrada.");
      const requestData = requestSnap.data() as MaterialRequest;

      if (requestData.status === 'ENTREGUE') {
        showToast("Esta solicitaÃ§Ã£o jÃ¡ foi entregue.", "info");
        return;
      }

      // Pre-fetch all necessary stock data with normalized name matching
      const itemsSnapshot = await getDocs(collection(db, 'items'));
      const allActiveItems = itemsSnapshot.docs
        .map(d => ({ id: d.id, ...d.data() } as Item))
        .filter(i => !i.deletedAt);

      const itemsStockData: any[] = [];
      for (const reqItem of requestItems) {
        if (reqItem.quantity_approved <= 0) continue;

        const normalizedReqName = normalizeString(reqItem.product_name);
        
        // Find all batches that represent this product (same normalized name)
        let batches = allActiveItems.filter(item => 
          normalizeString(item.name) === normalizedReqName && (item.quantity || 0) > 0
        );

        batches.sort((a, b) => {
          if (a.expiry_date === 'Indeterminada' || !a.expiry_date) return 1;
          if (b.expiry_date === 'Indeterminada' || !b.expiry_date) return -1;
          return new Date(a.expiry_date).getTime() - new Date(b.expiry_date).getTime();
        });

        let pharmItems: any[] = [];
        if (requestData.sector === 'FarmÃ¡cia') {
          pharmItems = allActiveItems
            .filter(item => normalizeString(item.name) === normalizedReqName && item.location === 'FarmÃ¡cia')
            .map(item => ({ id: item.id, batch_number: item.batch_number, ref: doc(db, 'items', item.id) }));
        }

        itemsStockData.push({ reqItem, batches, pharmItems });
      }

      await runTransaction(db, async (transaction) => {
        // Collect all batch and pharmacy refs to read them all first
        const batchRefs = itemsStockData.flatMap(d => d.batches.map(b => doc(db, 'items', b.id)));
        const pharmRefs = itemsStockData.flatMap(d => d.pharmItems.map(p => p.ref));
        
        // 1. Perform ALL reads first
        const [tRequestSnap, ...itemSnaps] = await Promise.all([
          transaction.get(requestRef),
          ...batchRefs.map(ref => transaction.get(ref)),
          ...pharmRefs.map(ref => transaction.get(ref))
        ]);

        const tRequestData = tRequestSnap.data() as MaterialRequest | undefined;
        if (!tRequestData || tRequestData.status === 'ENTREGUE') return;

        // Map snapshots for easy access by path
        const snapMap = new Map();
        itemSnaps.forEach(snap => snapMap.set(snap.ref.path, snap));

        // 2. Perform ALL writes
        transaction.update(requestRef, { 
          status: 'ENTREGUE',
          deliveredAt: new Date().toISOString(),
          deliveredBy: user?.email,
          updatedAt: serverTimestamp()
        });

        for (const { reqItem, batches, pharmItems } of itemsStockData) {
          let remaining = reqItem.quantity_approved;
          
          for (const batch of batches) {
            if (remaining <= 0) break;

            const tBatchRef = doc(db, 'items', batch.id);
            const tBatchSnap = snapMap.get(tBatchRef.path);
            if (!tBatchSnap || !tBatchSnap.exists()) continue;
            
            const tBatchData = tBatchSnap.data() as Item;
            const currentQty = tBatchData.quantity || 0;
            if (currentQty <= 0) continue;

            const toTake = Math.min(currentQty, remaining);
            
            transaction.update(tBatchRef, {
              quantity: currentQty - toTake,
              updatedAt: serverTimestamp()
            });

            // Log Transaction
            const transRef = doc(collection(db, 'transactions'));
            transaction.set(transRef, {
              item_id: batch.id,
              item_name: reqItem.product_name,
              type: 'exit',
              origin: batch.origin || 'extra',
              quantity: toTake,
              sector: requestData.sector,
              location: batch.location || 'Almoxarifado',
              date: new Date().toISOString(),
              responsible: user?.displayName || user?.email,
              responsibleEmail: user?.email,
              exitReason: 'consumo',
              batch_number: batch.batch_number,
              expiry_date: batch.expiry_date
            });

            if (requestData.sector === 'FarmÃ¡cia' && batch.location !== 'FarmÃ¡cia') {
              const existingPharm = pharmItems.find((p: any) => p.batch_number === batch.batch_number);
              if (existingPharm) {
                const tPharmRef = existingPharm.ref;
                const tPharmSnap = snapMap.get(tPharmRef.path);
                const tPharmData = tPharmSnap?.data() as Item | undefined;
                transaction.update(tPharmRef, {
                  quantity: (tPharmData?.quantity || 0) + toTake,
                  updatedAt: serverTimestamp()
                });
              } else {
                const newItemRef = doc(collection(db, 'items'));
                transaction.set(newItemRef, {
                  name: reqItem.product_name,
                  description: batch.description || '',
                  category: batch.category || 'Outros',
                  supplier: batch.supplier || 'TransferÃªncia',
                  batch_number: batch.batch_number || '',
                  expiry_date: batch.expiry_date || 'Indeterminada',
                  initial_quantity: toTake,
                  quantity: toTake,
                  min_quantity: batch.min_quantity || 0,
                  unit_price: batch.unit_price || 0,
                  location: 'FarmÃ¡cia',
                  origin: batch.origin || 'extra',
                  date: new Date().toISOString(),
                  createdAt: serverTimestamp(),
                  updatedAt: serverTimestamp()
                });
              }
            }
            remaining -= toTake;
          }

          if (remaining > 0) {
            throw new Error(`Estoque insuficiente para "${reqItem.product_name}".`);
          }
        }
      });

      // Cleanup and UI updates
      showToast("Entrega confirmada e estoque baixado!", "success");
      setShowRequestDetailModal({ show: false });

      // Notifications
      const uSnap = await getDocs(query(collection(db, 'users'), where('email', '==', requestData.requesterEmail)));
      if (!uSnap.empty) {
        await createNotification(uSnap.docs[0].id, 'Entrega Realizada', `Sua solicitaÃ§Ã£o #${requestId.slice(-5).toUpperCase()} foi entregue.`, requestId);
      }

      // Stock Zero Notifications
      for (const { reqItem } of itemsStockData) {
        await checkStockAndNotify(reqItem.product_name);
      }

      // Receipt
      const itemsForReceipt: { product_name: string; quantity: number; batch_number: string }[] = [];
      for (const { reqItem, batches } of itemsStockData) {
        let remaining = reqItem.quantity_approved;
        for (const batch of batches) {
          if (remaining <= 0) break;
          const available = batch.quantity || 0;
          if (available <= 0) continue;
          const toTake = Math.min(available, remaining);
          itemsForReceipt.push({
            product_name: reqItem.product_name,
            quantity: toTake,
            batch_number: batch.batch_number || '---'
          });
          remaining -= toTake;
        }
        if (remaining > 0 && !itemsForReceipt.some(i => i.product_name === reqItem.product_name)) {
          itemsForReceipt.push({
            product_name: reqItem.product_name,
            quantity: reqItem.quantity_approved,
            batch_number: '---'
          });
        }
      }
      handleExportDeliveryReceiptPDF({
        sector: requestData.sector,
        items: itemsForReceipt,
        requestId: requestId,
        date: new Date().toISOString()
      });

    } catch (error: any) {
      console.error("Erro na entrega:", error);
      showToast(`Erro: ${error.message}`, "error");
    }
  };

  const handleRejectRequest = async (requestId: string) => {
    try {
      await updateDoc(doc(db, 'requests', requestId), { 
        status: 'RECUSADO',
        adminObservation: adminObservation
      });
      
      const request = requests.find(r => r.id === requestId);
      if (request) {
        const userSnap = await getDocs(query(collection(db, 'users'), where('email', '==', request.requesterEmail)));
        if (!userSnap.empty) {
          const msg = adminObservation 
            ? `Sua solicitaÃ§Ã£o #${requestId.slice(-5).toUpperCase()} foi recusada. Motivo: ${adminObservation}`
            : `Sua solicitaÃ§Ã£o #${requestId.slice(-5).toUpperCase()} foi recusada.`;
          await createNotification(userSnap.docs[0].id, 'SolicitaÃ§Ã£o Recusada', msg, requestId);
        }
      }
      
      showToast("SolicitaÃ§Ã£o recusada.", "success");
    } catch (error: any) {
      handleFirestoreError(error, OperationType.UPDATE, `requests/${requestId}`);
      showToast(`Erro ao recusar: ${error.message}`, "error");
    }
  };

  const handleAddItem = async (e: React.FormEvent) => {
    e.preventDefault();
    
    try {
      for (const itemData of bulkEntry.items) {
        const trimmedName = itemData.name.trim();
        if (!trimmedName) {
          showToast("O nome do produto nÃ£o pode estar vazio ou conter apenas espaÃ§os.", "error");
          return;
        }

        const initial_qty = isNaN(itemData.initial_quantity) ? 0 : itemData.initial_quantity;
        
        // Dynamic min stock calculation (5 weeks coverage)
        const weeklyRate = weeklyExitRates[trimmedName] || 0;
        const calculatedMin = weeklyRate > 0 ? Math.ceil(weeklyRate * 5) : 5;
        const min_qty = isNaN(itemData.min_quantity) ? calculatedMin : itemData.min_quantity;
        
        // Inherit price from existing batches if not provided
        const existingPrice = items.find(i => i.name.toLowerCase() === trimmedName.toLowerCase() && (Number(i.unit_price) || 0) > 0)?.unit_price || 0;
        const price = isNaN(itemData.unit_price) || itemData.unit_price === 0 ? existingPrice : itemData.unit_price;

        // Check if item already exists with the same name AND batch AND location
        const existingItem = items.find(i => 
          i.name.toLowerCase() === trimmedName.toLowerCase() && 
          (i.batch_number || '').toLowerCase() === (itemData.batch_number || '').toLowerCase() &&
          (i.location || 'Almoxarifado') === inventoryLocation
        );

        if (existingItem) {
          await runTransaction(db, async (transaction) => {
            const itemDoc = doc(db, 'items', existingItem.id);
            const itemSnap = await transaction.get(itemDoc);
            
            if (!itemSnap.exists()) {
              throw new Error("Item nÃ£o encontrado durante a atualizaÃ§Ã£o.");
            }
            
            const currentItemData = itemSnap.data() as Item;
            const transCol = collection(db, 'transactions');
            
            const expiryValue = itemData.is_indeterminate_expiry ? 'Indeterminada' : itemData.expiry_date;

            transaction.update(itemDoc, {
              quantity: (Number(currentItemData.quantity) || 0) + initial_qty,
              min_quantity: min_qty,
              expiry_date: expiryValue || currentItemData.expiry_date,
              unit_price: price || currentItemData.unit_price,
              unit_measure: itemData.unit_measure || currentItemData.unit_measure || 'Unidade (UN)',
              supplier: bulkEntry.supplier || currentItemData.supplier,
              category: bulkEntry.category || currentItemData.category,
              medication_type: bulkEntry.category === 'Medicamentos' ? (itemData.medication_type || currentItemData.medication_type || '') : '',
              room: bulkEntry.room || currentItemData.room,
              updatedAt: serverTimestamp()
            });

            const newTransRef = doc(transCol);
            transaction.set(newTransRef, {
              item_id: existingItem.id,
              item_name: existingItem.name,
              type: 'entry',
              origin: bulkEntry.origin,
              quantity: initial_qty,
              location: inventoryLocation,
              room: bulkEntry.room,
              date: new Date().toISOString(),
              responsible: user?.displayName || 'Sistema',
              responsibleEmail: user?.email || '',
              supplier: bulkEntry.supplier || currentItemData.supplier,
              batch_number: itemData.batch_number,
              expiry_date: expiryValue,
              medication_type: bulkEntry.category === 'Medicamentos' ? (itemData.medication_type || '') : ''
            });
          });
        } else {
          const itemCol = collection(db, 'items');
          const transCol = collection(db, 'transactions');
          
          const expiryValue = itemData.is_indeterminate_expiry ? 'Indeterminada' : itemData.expiry_date;

          const itemRef = await addDoc(itemCol, {
            name: trimmedName,
            min_quantity: min_qty,
            expiry_date: expiryValue,
            origin: bulkEntry.origin,
            unit_price: price,
            unit_measure: itemData.unit_measure || 'Unidade (UN)',
            supplier: bulkEntry.supplier,
            category: bulkEntry.category,
            medication_type: bulkEntry.category === 'Medicamentos' ? (itemData.medication_type || '') : '',
            room: bulkEntry.room,
            batch_number: itemData.batch_number,
            quantity: initial_qty,
            location: inventoryLocation,
            createdAt: new Date().toISOString()
          });

          await addDoc(transCol, {
            item_id: itemRef.id,
            item_name: trimmedName,
            type: 'entry',
            origin: bulkEntry.origin,
            quantity: initial_qty,
            location: inventoryLocation,
            room: bulkEntry.room,
            date: new Date().toISOString(),
            responsible: user?.displayName || 'Sistema',
            responsibleEmail: user?.email || '',
            supplier: bulkEntry.supplier,
            batch_number: itemData.batch_number,
            expiry_date: expiryValue,
            medication_type: bulkEntry.category === 'Medicamentos' ? (itemData.medication_type || '') : ''
          });
        }
      }

      setShowAddModal(false);
      setBulkEntry({ 
        supplier: '',
        category: 'Expediente',
        origin: 'extra',
        room: 'Almoxarifado Principal',
        items: [{
          id: Math.random().toString(36).substr(2, 9),
          name: '',
          initial_quantity: 1,
          min_quantity: NaN,
          batch_number: '',
          expiry_date: '',
          is_indeterminate_expiry: false,
          unit_price: 0,
          unit_measure: 'Unidade (UN)',
          medication_type: ''
        }]
      });
    } catch (error: any) {
      console.error('Erro ao salvar itens:', error);
      alert(`Erro ao salvar itens: ${error.message}`);
    }
  };

  const handleTransaction = async (e: React.FormEvent) => {
    e.preventDefault();
    
    try {
      if (showTransactionModal.type === 'exit') {
        if (basket.length === 0) return;
        
        await runTransaction(db, async (transaction) => {
          const processedItems = [];
          
          for (const b of basket) {
            const itemRef = doc(db, 'items', b.item_id);
            const itemSnap = await transaction.get(itemRef);
            
            if (!itemSnap.exists()) {
              throw new Error(`Item ${b.item_id} nÃ£o encontrado.`);
            }

            const currentItemData = itemSnap.data() as Item;
            const currentQty = Number(currentItemData.quantity) || 0;
            if (currentQty < b.quantity) {
              throw new Error(`Estoque insuficiente para o item ${currentItemData.name}. DisponÃ­vel: ${currentQty}`);
            }

            let pharmacyItemSnap = null;
            if (selectedSector === 'FarmÃ¡cia' && exitReason === 'consumo') {
              const pharmacyItemsQuery = query(
                collection(db, 'items'),
                where('name', '==', currentItemData.name),
                where('batch_number', '==', currentItemData.batch_number || ''),
                where('location', '==', 'FarmÃ¡cia')
              );
              const fullSnap = await getDocs(pharmacyItemsQuery);
              const activeDocs = fullSnap.docs.filter(d => !d.data().deletedAt);
              pharmacyItemSnap = {
                empty: activeDocs.length === 0,
                docs: activeDocs
              };
            }

            processedItems.push({
              itemRef,
              currentItemData,
              quantity: b.quantity,
              pharmacyItemSnap
            });
          }

          const transCol = collection(db, 'transactions');
          const itemsCol = collection(db, 'items');

          for (const pi of processedItems) {
            const { itemRef, currentItemData, quantity, pharmacyItemSnap } = pi;
            const currentQty = Number(currentItemData.quantity) || 0;

            transaction.update(itemRef, {
              quantity: currentQty - quantity,
              updatedAt: serverTimestamp()
            });

            const newTransRef = doc(transCol);
            const currentDonationNumber = exitReason === 'doacao' ? (() => {
              const currentYear = new Date().getFullYear();
              const yearlyDonations = transactions.filter(t => 
                t.exitReason === 'doacao' && 
                !t.deletedAt && 
                new Date(t.date).getFullYear() === currentYear
              );
              const uniqueDonations = new Set();
              yearlyDonations.forEach(t => {
                // Group by either donationNumber or a "session key" (rough timestamp + destinatario)
                if ((t as any).donationNumber) {
                  uniqueDonations.add((t as any).donationNumber);
                } else {
                  // Fallback for older transactions: group by date (minute precision) and sector
                  const dateKey = new Date(t.date).toISOString().slice(0, 16);
                  uniqueDonations.add(`${dateKey}-${t.sector}`);
                }
              });
              const nextCount = uniqueDonations.size + 1;
              return `${nextCount.toString().padStart(2, '0')}/${currentYear}`;
            })() : null;

            const sectorValue = modalSector || (inventoryLocation === 'FarmÃ¡cia' ? 'FarmÃ¡cia (Consumo Interno)' : 'Almoxarifado');

            transaction.set(newTransRef, {
              item_id: currentItemData.id || itemRef.id,
              item_name: currentItemData.name,
              type: 'exit',
              origin: currentItemData.origin,
              quantity: quantity,
              sector: sectorValue,
              location: inventoryLocation,
              date: new Date().toISOString(),
              responsible: user?.displayName || 'Sistema',
              responsibleEmail: user?.email || '',
              exitReason: exitReason,
              expiryReason: (exitReason === 'vencido' || exitReason === 'perda') ? expiryReason : null,
              donationUnitName: exitReason === 'doacao' ? (donationUnitName || 'CEO - Centro de Especialidades OdontolÃ³gicas') : null,
              donationUnitAddress: exitReason === 'doacao' ? donationUnitAddress : null,
              donationUnitCNPJ: exitReason === 'doacao' ? donationUnitCNPJ : null,
              donationRevisionDate: exitReason === 'doacao' ? donationRevisionDate : null,
              donationNumber: currentDonationNumber,
              batch_number: currentItemData.batch_number,
              expiry_date: currentItemData.expiry_date
            });

            if (pharmacyItemSnap) {
              let pharmacyItemId = '';
              if (!pharmacyItemSnap.empty) {
                const pharmacyItemDoc = pharmacyItemSnap.docs[0];
                pharmacyItemId = pharmacyItemDoc.id;
                transaction.update(pharmacyItemDoc.ref, {
                  quantity: (pharmacyItemDoc.data().quantity || 0) + quantity,
                  updatedAt: serverTimestamp()
                });
              } else {
                const newItemRef = doc(itemsCol);
                pharmacyItemId = newItemRef.id;
                transaction.set(newItemRef, {
                  name: currentItemData.name,
                  description: currentItemData.description || '',
                  quantity: quantity,
                  min_quantity: currentItemData.min_quantity || 5,
                  expiry_date: currentItemData.expiry_date,
                  origin: currentItemData.origin,
                  unit_price: currentItemData.unit_price,
                  supplier: currentItemData.supplier,
                  category: currentItemData.category,
                  batch_number: currentItemData.batch_number,
                  location: 'FarmÃ¡cia',
                  createdAt: new Date().toISOString()
                });
              }

              const pharmTransRef = doc(transCol);
              transaction.set(pharmTransRef, {
                item_id: pharmacyItemId,
                item_name: currentItemData.name,
                type: 'entry',
                origin: currentItemData.origin,
                quantity: quantity,
                location: 'FarmÃ¡cia',
                date: new Date().toISOString(),
                responsible: 'Sistema (TransferÃªncia)',
                batch_number: currentItemData.batch_number,
                expiry_date: currentItemData.expiry_date,
                supplier: currentItemData.supplier
              });
            }
          }
        });
      } else {
        const item = showTransactionModal.item || items.find(i => i.id === selectedItemId);
        if (!item) {
          alert('Por favor, selecione um item.');
          return;
        }

        const weeklyRate = weeklyExitRates[item.name] || 0;
        const calculatedMin = weeklyRate > 0 ? Math.ceil(weeklyRate * 5) : item.min_quantity;
        const finalMinStock = isNaN(transactionMinStock) ? calculatedMin : transactionMinStock;
        
        await runTransaction(db, async (transaction) => {
          const itemDoc = doc(db, 'items', item.id);
          const itemSnap = await transaction.get(itemDoc);
          
          if (!itemSnap.exists()) {
            throw new Error("Item nÃ£o encontrado.");
          }

          const currentItemData = itemSnap.data() as Item;
          const transCol = collection(db, 'transactions');
          
          transaction.update(itemDoc, {
            quantity: (Number(currentItemData.quantity) || 0) + transactionQty,
            min_quantity: finalMinStock,
            updatedAt: serverTimestamp()
          });

          const newTransRef = doc(transCol);
          transaction.set(newTransRef, {
            item_id: item.id,
            item_name: currentItemData.name,
            type: 'entry',
            origin: currentItemData.origin,
            quantity: transactionQty,
            sector: null,
            location: inventoryLocation,
            date: new Date().toISOString(),
            responsible: user?.displayName || 'Sistema',
            responsibleEmail: user?.email || '',
            batch_number: currentItemData.batch_number,
            expiry_date: currentItemData.expiry_date,
            supplier: currentItemData.supplier
          });
        });
      }

      setShowTransactionModal({ show: false, type: 'entry' });
      
      // Auto-generate delivery receipt for manual exit
      if (showTransactionModal.type === 'exit' && basket.length > 0 && selectedSector) {
        const itemsForReceipt = basket.map(b => {
          const foundItem = items.find(i => i.id === b.item_id);
          return {
            product_name: foundItem?.name || 'Produto NÃ£o Identificado',
            quantity: b.quantity,
            batch_number: foundItem?.batch_number || '---'
          };
        });
        
        if (exitReason === 'doacao') {
          // Calculate donation number for this year
          const currentYear = new Date().getFullYear();
          const yearlyDonations = transactions.filter(t => 
            t.exitReason === 'doacao' && 
            !t.deletedAt && 
            new Date(t.date).getFullYear() === currentYear
          );
          const uniqueDonations = new Set();
          yearlyDonations.forEach(t => {
            if ((t as any).donationNumber) {
              uniqueDonations.add((t as any).donationNumber);
            } else {
              const dateKey = new Date(t.date).toISOString().slice(0, 16);
              uniqueDonations.add(`${dateKey}-${t.sector}`);
            }
          });
          const currentDonationNumber = `${(uniqueDonations.size + 1).toString().padStart(2, '0')}/${currentYear}`;

          handleExportDonationTermPDF({
            donatingUnitName: donationUnitName || 'CEO - Centro de Especialidades OdontolÃ³gicas',
            receivingUnit: {
              name: selectedSector || 'Unidade Receptora',
              address: donationUnitAddress,
              cnpj: donationUnitCNPJ
            },
            items: itemsForReceipt,
            revisionDate: donationRevisionDate,
            donationNumber: currentDonationNumber,
            date: new Date().toISOString()
          });
        } else {
          handleExportDeliveryReceiptPDF({
            sector: selectedSector,
            items: itemsForReceipt,
            date: new Date().toISOString()
          });
        }
      }

      setTransactionMinStock(NaN);
      setTransactionQty(1);
      setExitReason('consumo');
      setExpiryReason('');
      setSelectedSector(SECTORS[0]);
      setSelectedItemId('');
      setBasket([]);
      setDonationUnitName('');
      setDonationUnitAddress('');
      setDonationUnitCNPJ('');
      setDonationRevisionDate('');
      setLetterheadImage(null);

      // Stock Zero Notifications check
      if (showTransactionModal.type === 'exit') {
        const itemNames = basket.map(b => items.find(i => i.id === b.item_id)?.name).filter(Boolean) as string[];
        for (const name of itemNames) {
          await checkStockAndNotify(name);
        }
      }
    } catch (error: any) {
      console.error('Erro na transaÃ§Ã£o:', error);
      alert(`Erro na movimentaÃ§Ã£o: ${error.message}`);
    }
  };

  const handleExportExcel = () => {
    try {
      // Prepare data for Excel
      const exportData: any[] = [];
      reportData.consumptionReport.forEach(item => {
        // Main item row
        const row: any = {
          'Item': item.name,
          'Categoria': item.category,
          'Fornecedor': item.supplier,
          'Quantidade Total': item.totalQuantity,
          'Destino': 'TOTAL'
        };
        if (isAdmin) row['Valor Total (BRL)'] = item.totalValue;
        exportData.push(row);
        
        // Sector breakdown rows
        Object.entries(item.sectors).forEach(([sector, qty]) => {
          const subRow: any = {
            'Item': `   -> ${item.name}`,
            'Categoria': item.category,
            'Fornecedor': item.supplier,
            'Quantidade Total': qty,
            'Destino': sector
          };
          if (isAdmin) subRow['Valor Total (BRL)'] = '';
          exportData.push(subRow);
        });
      });

      // Create worksheet
      const ws = XLSX.utils.json_to_sheet(exportData);
      
      // Create workbook
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "RelatÃ³rio de SaÃ­das");

      // Generate filename
      const dateStr = format(new Date(), 'dd-MM-yyyy');
      const sectorStr = reportSectorFilter === 'all' ? 'Todos_Setores' : reportSectorFilter.replace(/\s+/g, '_');
      const fileName = `Relatorio_Estoque_${sectorStr}_${dateStr}.xlsx`;

      // Save file
      XLSX.writeFile(wb, fileName);
    } catch (error) {
      console.error('Erro ao exportar Excel:', error);
      alert('Ocorreu um erro ao gerar o arquivo Excel.');
    }
  };

  const handleExportInventory = () => {
    try {
      const exportData = groupedArray.map(group => {
        let status = group.total_quantity <= group.min_quantity ? 'BAIXO' : 'OK';
        if (group.durationWeeks !== 'infinite') {
          if (group.durationWeeks <= 4) status = 'MUITO CRÃTICO';
          else if (group.durationWeeks <= 8) status = 'CRÃTICO';
        }
        
        return {
          'Item': group.name,
          'Categoria': group.category || '---',
          'Estoque Total': group.total_quantity,
          'MÃ­nimo': group.min_quantity,
          'DuraÃ§Ã£o (Semanas)': group.durationWeeks === 'infinite' ? 'Infinito' : group.durationWeeks.toFixed(1),
          'Status': status
        };
      });

      const ws = XLSX.utils.json_to_sheet(exportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Estoque Atual");
      const dateStr = format(new Date(), 'dd-MM-yyyy');
      XLSX.writeFile(wb, `Estoque_Atual_${dateStr}.xlsx`);
      showToast("Estoque exportado com sucesso!", "success");
    } catch (error) {
      console.error('Erro ao exportar estoque:', error);
      showToast("Erro ao exportar estoque.", "error");
    }
  };

  const handleExportInventoryPDF = () => {
    try {
      const doc = new jsPDF();
      const startY = drawPDFLetterhead(
        doc,
        'RelatÃ³rio de Estoque Atual',
        `Gerado em: ${format(new Date(), 'dd/MM/yyyy HH:mm')}`
      );

      // Prepare data for table
      const tableData = groupedArray.map(group => {
        let status = group.total_quantity <= group.min_quantity ? 'BAIXO' : 'OK';
        if (group.durationWeeks !== 'infinite') {
          if (group.durationWeeks <= 4) status = 'MUITO CRÃTICO';
          else if (group.durationWeeks <= 8) status = 'CRÃTICO';
        }
        
        return [
          group.name,
          group.category || '---',
          group.total_quantity.toString(),
          group.durationWeeks === 'infinite' ? 'Infinito' : group.durationWeeks.toFixed(1),
          group.min_quantity.toString(),
          status
        ];
      });
      
      // Generate table
      autoTable(doc, {
        startY: startY + 4,
        head: [['Item', 'Categoria', 'Estoque', 'DuraÃ§Ã£o (Sem)', 'MÃ­nimo', 'Status']],
        body: tableData,
        theme: 'striped',
        headStyles: { fillColor: [28, 25, 23], halign: 'center' }, // #1C1917
        columnStyles: {
          2: { halign: 'center' },
          3: { halign: 'center' },
          4: { halign: 'center' },
          5: { halign: 'center' }
        },
        styles: { fontSize: 9, cellPadding: 3 },
        didParseCell: function(data) {
          if (data.section === 'body' && data.column.index === 5) {
            const text = data.cell.text[0];
            if (text === 'BAIXO' || text === 'MUITO CRÃTICO') {
              data.cell.styles.textColor = [225, 29, 72]; // rose-600
              data.cell.styles.fontStyle = 'bold';
            } else if (text === 'CRÃTICO') {
              data.cell.styles.textColor = [249, 115, 22]; // orange-500
              data.cell.styles.fontStyle = 'bold';
            }
          }
        }
      });
      
      // Save PDF
      const dateStr = format(new Date(), 'dd-MM-yyyy');
      doc.save(`Estoque_Atual_${dateStr}.pdf`);
      showToast("PDF de estoque exportado com sucesso!", "success");
    } catch (error) {
      console.error('Erro ao exportar PDF de estoque:', error);
      showToast("Erro ao exportar PDF de estoque.", "error");
    }
  };

  const handleExportLowStockPDF = () => {
    try {
      const doc = new jsPDF();
      const dateStr = format(new Date(), 'dd/MM/yyyy HH:mm');
      const locationLabel = inventoryLocation === 'FarmÃ¡cia' ? 'FarmÃ¡cia (Medicamentos)' : 'Almoxarifado Geral';
      
      const startY = drawPDFLetterhead(
        doc,
        `RELATÃ“RIO DE ITENS CRÃTICOS - ESTOQUE BAIXO`,
        `Unidade: ${locationLabel} - Data de EmissÃ£o: ${dateStr}`
      );

      // Collect all active items for current location
      const activeLocationItems = items.filter(
        i => !i.deletedAt && i.quantity > 0 && (i.location || 'Almoxarifado') === inventoryLocation
      );

      const locationGrouped: Record<string, ItemGroup> = {};
      activeLocationItems.forEach(item => {
        if (!locationGrouped[item.name]) {
          const weeklyRate = weeklyExitRates[item.name] || 0;
          locationGrouped[item.name] = {
            name: item.name,
            total_quantity: 0,
            min_quantity: weeklyRate > 0 ? Math.ceil(weeklyRate * 5) : item.min_quantity,
            category: item.category,
            supplier: item.supplier,
            unit_measure: item.unit_measure || null,
            batches: [],
            weeklyExitRate: weeklyRate,
            durationWeeks: 0
          };
        }
        locationGrouped[item.name].total_quantity += item.quantity;
        if (!locationGrouped[item.name].unit_measure && item.unit_measure) {
          locationGrouped[item.name].unit_measure = item.unit_measure;
        }
        locationGrouped[item.name].batches.push(item);
      });

      // Filter groups where total_quantity <= min_quantity
      const lowStockGroupsList = Object.values(locationGrouped).filter(
        g => g.total_quantity <= g.min_quantity
      );

      // Sort by deficit/ratio ascending (most critical first)
      lowStockGroupsList.sort((a, b) => {
        const ratioA = a.min_quantity > 0 ? a.total_quantity / a.min_quantity : 1;
        const ratioB = b.min_quantity > 0 ? b.total_quantity / b.min_quantity : 1;
        return ratioA - ratioB;
      });

      if (lowStockGroupsList.length === 0) {
        doc.setFontSize(11);
        doc.setTextColor(16, 185, 129); // emerald-600
        doc.setFont("helvetica", "bold");
        doc.text("Nenhum item com estoque baixo registrado no momento.", 14, startY + 12);
        
        doc.setFontSize(9.5);
        doc.setTextColor(100, 116, 139);
        doc.setFont("helvetica", "normal");
        doc.text(`Todos os insumos cadastrados no ${locationLabel} estÃ£o acima do estoque mÃ­nimo.`, 14, startY + 20);

        const dateFileStr = format(new Date(), 'dd-MM-yyyy');
        doc.save(`Relatorio_Estoque_Baixo_${inventoryLocation}_${dateFileStr}.pdf`);
        showToast("RelatÃ³rio gerado: Nenhum item com estoque baixo encontrado.", "info");
        return;
      }

      // Prepare table data
      const tableData = lowStockGroupsList.map(group => {
        const deficit = Math.max(0, group.min_quantity - group.total_quantity);
        let status = 'ESTOQUE BAIXO';
        if (group.total_quantity === 0) {
          status = 'ZERADO / SEM ESTOQUE';
        } else if (group.total_quantity <= (group.min_quantity * 0.5)) {
          status = 'MUITO CRÃTICO';
        }

        const unitText = group.unit_measure ? ` ${group.unit_measure}` : '';

        return [
          group.name,
          group.category || 'Geral',
          `${group.total_quantity}${unitText}`,
          `${group.min_quantity}${unitText}`,
          `${deficit}${unitText}`,
          status
        ];
      });

      autoTable(doc, {
        startY: startY + 4,
        head: [['Material / Medicamento', 'Categoria', 'Estoque Atual', 'Estoque MÃ­nimo', 'DÃ©ficit (ReposiÃ§Ã£o)', 'Status CrÃ­tico']],
        body: tableData,
        theme: 'striped',
        headStyles: { fillColor: [180, 35, 24], halign: 'center', fontStyle: 'bold' }, // Dark Red/Amber
        columnStyles: {
          2: { halign: 'center' },
          3: { halign: 'center' },
          4: { halign: 'center', fontStyle: 'bold' },
          5: { halign: 'center' }
        },
        styles: { fontSize: 8.5, cellPadding: 3 },
        didParseCell: function(data) {
          if (data.section === 'body' && data.column.index === 5) {
            const text = data.cell.text[0];
            if (text.includes('ZERADO') || text === 'MUITO CRÃTICO') {
              data.cell.styles.textColor = [220, 38, 38]; // red-600
              data.cell.styles.fontStyle = 'bold';
            } else if (text === 'ESTOQUE BAIXO') {
              data.cell.styles.textColor = [217, 119, 6]; // amber-600
              data.cell.styles.fontStyle = 'bold';
            }
          }
        }
      });

      // Add total count summary at bottom
      const finalY = (doc as any).lastAutoTable ? (doc as any).lastAutoTable.finalY + 10 : startY + 50;
      doc.setFontSize(9);
      doc.setTextColor(30, 41, 59);
      doc.setFont("helvetica", "bold");
      doc.text(`Total de itens identificados com estoque baixo ou crÃ­tico: ${lowStockGroupsList.length}`, 14, finalY);

      const dateFileStr = format(new Date(), 'dd-MM-yyyy');
      doc.save(`Relatorio_Estoque_Baixo_${inventoryLocation}_${dateFileStr}.pdf`);
      showToast("RelatÃ³rio PDF de estoque baixo gerado com sucesso!", "success");
    } catch (error) {
      console.error('Erro ao exportar PDF de estoque baixo:', error);
      showToast("Erro ao exportar PDF de estoque baixo.", "error");
    }
  };

  const handleExportMaterialsCatalogPDF = () => {
    try {
      // @ts-ignore
      const doc = new jsPDF();
      const pageWidth = doc.internal.pageSize.width;
      
      const startY = drawPDFLetterhead(
        doc,
        'CatÃ¡logo de Materiais em Estoque',
        `CEO - Centro de Especialidades OdontolÃ³gicas - Gerado em: ${format(new Date(), 'dd/MM/yyyy HH:mm')}`
      );
      
      // Filter unique items across all batches and locations
      const uniqueItems: Record<string, { name: string, category: string, supplier: string }> = {};
      
      items.filter(i => !i.deletedAt && i.quantity > 0).forEach(item => {
        if (!uniqueItems[item.name]) {
          uniqueItems[item.name] = {
            name: item.name,
            category: item.category || '---',
            supplier: item.supplier || '---'
          };
        }
      });
      
      const tableData = Object.values(uniqueItems)
        .sort((a, b) => a.name.localeCompare(b.name))
        .map(item => [
          item.name,
          item.category,
          item.supplier
        ]);
      
      // Generate table (NO Stock, NO Batch, NO Expiry)
      autoTable(doc, {
        startY: startY + 4,
        head: [['Material / Produto', 'Categoria', 'Fornecedor']],
        body: tableData,
        theme: 'striped',
        headStyles: { fillColor: [28, 25, 23], halign: 'left' }, // #1C1917
        styles: { fontSize: 9, cellPadding: 3.5 },
        columnStyles: {
          0: { cellWidth: 'auto' },
          1: { cellWidth: 50 },
          2: { cellWidth: 50 }
        }
      });
      
      // Footer on every page
      const pageCount = (doc as any).internal.getNumberOfPages();
      for (let i = 1; i <= pageCount; i++) {
        doc.setPage(i);
        doc.setFontSize(8);
        doc.setTextColor(168, 162, 158);
        doc.text(
          `PÃ¡gina ${i} de ${pageCount} - CatÃ¡logo gerado para consulta administrativa`,
          pageWidth / 2,
          doc.internal.pageSize.height - 10,
          { align: 'center' }
        );
      }
      
      const dateStr = format(new Date(), 'dd-MM-yyyy');
      doc.save(`Catalogo_Materiais_${dateStr}.pdf`);
      showToast("CatÃ¡logo de materiais exportado com sucesso!", "success");
    } catch (error) {
      console.error('Erro ao exportar catÃ¡logo:', error);
      showToast("Erro ao exportar catÃ¡logo de materiais.", "error");
    }
  };

  const handleExportRequestsPDF = () => {
    try {
      const doc = new jsPDF();
      
      const startY = drawPDFLetterhead(
        doc,
        'RelatÃ³rio de SolicitaÃ§Ãµes de Materiais',
        `Gerado em: ${format(new Date(), 'dd/MM/yyyy HH:mm')}`
      );
      
      // Determine which requests to export based on current tab
      let requestsToExport = [];
      if (activeTab === 'requests') {
        requestsToExport = requests.filter(req => !req.deletedAt && !req.isReturn);
      } else if (activeTab === 'admin-devolutions') {
        requestsToExport = requests.filter(req => !req.deletedAt && req.isReturn);
      } else if (activeTab === 'my-requests') {
        requestsToExport = requests.filter(r => r.sector === selectedSector && !r.deletedAt && !r.isReturn);
      } else if (activeTab === 'devolution') {
        requestsToExport = requests.filter(r => r.sector === selectedSector && !r.deletedAt && r.isReturn);
      } else {
        requestsToExport = requests.filter(req => !req.deletedAt);
      }
      
      // Sort by date descending
      requestsToExport.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());

      // Prepare data for table
      const tableData = requestsToExport.map(req => [
        `#${req.id.slice(-5).toUpperCase()}`,
        format(new Date(req.date), 'dd/MM/yyyy'),
        req.sector,
        req.status,
        allRequestItems.filter(ri => ri.request_id === req.id).length.toString()
      ]);
      
      // Generate table
      autoTable(doc, {
        startY: startY + 4,
        head: [['NÂº', 'Data', 'Setor', 'Status', 'Itens']],
        body: tableData,
        theme: 'striped',
        headStyles: { fillColor: [28, 25, 23], halign: 'center' }, // #1C1917
        columnStyles: {
          0: { halign: 'center' },
          1: { halign: 'center' },
          3: { halign: 'center' },
          4: { halign: 'center' }
        },
        styles: { fontSize: 9, cellPadding: 3 }
      });
      
      // Save PDF
      const fileName = `Solicitacoes_${format(new Date(), 'dd-MM-yyyy')}.pdf`;
      doc.save(fileName);
      showToast("PDF exportado com sucesso!", "success");
    } catch (error) {
      console.error('Erro ao exportar PDF:', error);
      showToast("Erro ao exportar PDF.", "error");
    }
  };

  useEffect(() => {
    if (showTransactionModal.show) {
      setTransactionQty(1);
      setExitReason('consumo');
      setExpiryReason('');
      setDonationUnitName('');
      setDonationUnitAddress('');
      setDonationUnitCNPJ('');
      setDonationRevisionDate('');
      setBasket(showTransactionModal.item ? [{ item_id: showTransactionModal.item.id!, quantity: 1 }] : []);
      
      // Default to item's current location or parent sector
      if (showTransactionModal.item) {
        setModalSector(showTransactionModal.item.location === 'FarmÃ¡cia' ? 'FarmÃ¡cia' : 'Almoxarifado');
      } else {
        setModalSector(userProfile?.sector || SECTORS[0]);
      }
    }
  }, [showTransactionModal.show, showTransactionModal.item, userProfile?.sector]);

  const handleExportPCA = () => {
    if (selectedSector !== 'Almoxarifado') {
      showToast("Acesso restrito ao Almoxarifado.", "error");
      return;
    }
    try {
      const doc = new jsPDF();
      const start = startOfDay(parseISO(pcaRange.start));
      const end = endOfDay(parseISO(pcaRange.end));

      const consumptionTransactions = transactions.filter(t => {
        if (t.deletedAt) return false;
        if (t.type !== 'exit' || t.exitReason !== 'consumo') return false;
        const d = new Date(t.date);
        return d >= start && d <= end;
      });

      // Grouping by category
      const groupedData: Record<string, Record<string, { name: string, quantity: number, unit: string }>> = {};

      consumptionTransactions.forEach(t => {
        const item = items.find(i => i.id === t.item_id);
        const category = item?.category || 'Outros';
        
        if (pcaCategory !== 'all' && category !== pcaCategory) return;

        if (!groupedData[category]) {
          groupedData[category] = {};
        }

        const itemName = t.item_name;
        if (!groupedData[category][itemName]) {
          groupedData[category][itemName] = {
            name: itemName,
            quantity: 0,
            unit: item?.description || 'UN'
          };
        }
        groupedData[category][itemName].quantity += t.quantity;
      });

      // Logo/Header
      let currentY = 20;
      if (letterheadImage) {
        try {
          doc.addImage(letterheadImage, 'PNG', 14, currentY, 182, 25);
          currentY += 30;
        } catch (e) {
          currentY += 5;
        }
      }

      doc.setFontSize(18);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(28, 25, 23);
      doc.text('RelatÃ³rio PCA - Plano Anual de ContrataÃ§Ã£o', 105, currentY, { align: 'center' });
      currentY += 8;
      
      doc.setFontSize(10);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(120, 113, 108);
      doc.text(`PerÃ­odo de Consumo: ${format(start, 'dd/MM/yyyy')} atÃ© ${format(end, 'dd/MM/yyyy')}`, 105, currentY, { align: 'center' });
      currentY += 5;
      doc.text(`Gerado em: ${format(new Date(), 'dd/MM/yyyy HH:mm')}`, 105, currentY, { align: 'center' });
      currentY += 15;

      const categories = Object.keys(groupedData).sort();
      
      if (categories.length === 0) {
        doc.setFontSize(12);
        doc.text('Nenhum consumo registrado no perÃ­odo selecionado.', 105, currentY + 20, { align: 'center' });
      } else {
        categories.forEach((category) => {
          const itemsInCategory = Object.values(groupedData[category]).sort((a, b) => a.name.localeCompare(b.name));
          
          const tableData = itemsInCategory.map(item => [
            item.name,
            `${item.quantity}`,
            item.unit
          ]);

          if (currentY > 230) {
            doc.addPage();
            currentY = 20;
          }

          // Category Header - Modern and Minimalist
          doc.setFontSize(11);
          doc.setFont('helvetica', 'bold');
          doc.setTextColor(30, 64, 175); // Dark blue text
          doc.setFillColor(239, 246, 255); // Very light blue background
          doc.rect(14, currentY, 182, 10, 'F');
          
          // Thin border for header
          doc.setDrawColor(191, 219, 254);
          doc.rect(14, currentY, 182, 10, 'S');
          
          doc.text(category.toUpperCase(), 18, currentY + 7);
          currentY += 12;

          autoTable(doc, {
            startY: currentY,
            head: [['Material', 'Quantidade Total Consumida', 'Unidade']],
            body: tableData,
            theme: 'grid',
            headStyles: { 
              fillColor: [248, 250, 252], 
              textColor: [71, 85, 105], 
              fontSize: 9, 
              fontStyle: 'bold',
              lineWidth: 0.1,
              lineColor: [226, 232, 240]
            },
            bodyStyles: { 
              fontSize: 8, 
              textColor: [30, 41, 59],
              lineWidth: 0.1,
              lineColor: [241, 245, 249]
            },
            alternateRowStyles: {
              fillColor: [250, 250, 250]
            },
            margin: { left: 14, right: 14 },
            styles: {
              cellPadding: 3
            }
          });

          currentY = (doc as any).lastAutoTable.finalY + 15;
        });
      }

      doc.save(`Relatorio_PCA_${format(new Date(), 'dd_MM_yyyy')}.pdf`);
      showToast("RelatÃ³rio PCA gerado com sucesso!", "success");
    } catch (error) {
      console.error('Erro ao gerar relatÃ³rio PCA:', error);
      showToast("Erro ao gerar relatÃ³rio PCA.", "error");
    }
  };

  const quantitativoReportData = useMemo(() => {
    if (quantitativoSource === 'sample') {
      return {
        months: ['Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'],
        monthColors: ['#1d4ed8', '#b91c1c', '#b45309', '#15803d', '#c2410c', '#0284c7'],
        sectors: [
          { name: 'ALMOXARIFADO', values: [10, 0, 0, 0, 0, 0], total: 10 },
          { name: 'CER', values: [11, 64, 19, 13, 27, 6], total: 140 },
          { name: 'CME', values: [4, 30, 0, 15, 4, 0], total: 53 },
          { name: 'ENVASE', values: [2, 0, 0, 5, 1, 1], total: 9 },
          { name: 'ESC. QUALIDADE', values: [80, 0, 0, 0, 0, 0], total: 80 },
          { name: 'HIGIENIZAÃ‡ÃƒO', values: [4, 0, 0, 1, 3, 0], total: 8 },
          { name: 'ILHA', values: [316, 178, 266, 310, 579, 200], total: 1849 },
          { name: 'IMAGEM', values: [351, 354, 131, 267, 505, 106], total: 1714 },
          { name: 'PÃ‰ DIABÃ‰TICO', values: [384, 476, 563, 548, 572, 552], total: 3095 },
          { name: 'RECEPÃ‡ÃƒO GERAL', values: [203, 0, 0, 0, 110, 17], total: 330 },
          { name: 'SINAIS VITAIS', values: [18, 8, 15, 8, 10, 9], total: 68 }
        ],
        title: quantitativoTitle,
        criticalAnalysis: quantitativoCriticalAnalysis
      };
    }

    let months: string[] = ['Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
    let monthColors = ['#1d4ed8', '#b91c1c', '#b45309', '#15803d', '#c2410c', '#0284c7', '#7c3aed', '#db2777', '#059669'];
    
    let startDate: Date;
    let endDate: Date;

    if (quantitativoPeriodPreset === '1_semestre_2026') {
      startDate = new Date('2026-01-01T00:00:00');
      endDate = new Date('2026-06-30T23:59:59');
      months = ['Janeiro', 'Fevereiro', 'MarÃ§o', 'Abril', 'Maio', 'Junho'];
    } else if (quantitativoPeriodPreset === '2_semestre_2026') {
      startDate = new Date('2026-07-01T00:00:00');
      endDate = new Date('2026-12-31T23:59:59');
      months = ['Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
    } else if (quantitativoPeriodPreset === 'ano_2026') {
      startDate = new Date('2026-01-01T00:00:00');
      endDate = new Date('2026-12-31T23:59:59');
      months = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
    } else {
      startDate = startOfDay(parseISO(quantitativoCustomStart));
      endDate = endOfDay(parseISO(quantitativoCustomEnd));
      months = ['Janeiro', 'Fevereiro', 'MarÃ§o', 'Abril', 'Maio', 'Junho'];
    }

    const sectorMap: Record<string, number[]> = {};

    SECTORS.forEach(sec => {
      sectorMap[sec.toUpperCase()] = new Array(months.length).fill(0);
    });

    const checkCategoryMatch = (itemCat: string | null | undefined, filterCat: string) => {
      if (!filterCat || filterCat === 'Todos' || filterCat.startsWith('Todos')) return true;
      if (!itemCat) return filterCat === 'Outros';

      const catLower = itemCat.toLowerCase().trim();
      const filterLower = filterCat.toLowerCase().trim();

      if (filterLower.includes('mÃ©dico') || filterLower.includes('medico') || filterLower.includes('hospitalar')) {
        return catLower.includes('mÃ©dico') || catLower.includes('medico') || catLower.includes('hospitalar');
      }
      if (filterLower.includes('medicamento')) {
        return catLower.includes('medicamento') || catLower.includes('fÃ¡rmaco') || catLower.includes('farmaco');
      }
      if (filterLower.includes('aliment')) {
        return catLower.includes('aliment') || catLower.includes('copa') || catLower.includes('cozinha');
      }
      if (filterLower.includes('expediente')) {
        return catLower.includes('expediente') || catLower.includes('papelaria') || catLower.includes('escritÃ³rio') || catLower.includes('escritorio');
      }
      if (filterLower.includes('higiene') || filterLower.includes('limpeza')) {
        return catLower.includes('higiene') || catLower.includes('limpeza') || catLower.includes('saneante');
      }
      if (filterLower.includes('odont')) {
        return catLower.includes('odont');
      }
      if (filterLower.includes('epi')) {
        return catLower.includes('epi') || catLower.includes('seguranÃ§a') || catLower.includes('seguranca');
      }
      if (filterLower.includes('informÃ¡t') || filterLower.includes('informat') || filterLower.includes('ti')) {
        return catLower.includes('informÃ¡t') || catLower.includes('informat') || catLower.includes('ti');
      }

      return catLower.includes(filterLower) || filterLower.includes(catLower);
    };

    transactions.forEach(t => {
      if (t.deletedAt) return;
      if (t.type !== 'exit') return;
      const tDate = new Date(t.date);
      if (tDate < startDate || tDate > endDate) return;

      const item = items.find(i => i.id === t.item_id);
      if (!checkCategoryMatch(item?.category, quantitativoCategory)) return;

      const secName = (t.sector || 'Outros').toUpperCase();
      if (!sectorMap[secName]) {
        sectorMap[secName] = new Array(months.length).fill(0);
      }

      let monthIdx = 0;
      if (months.length === 6) {
        monthIdx = tDate.getMonth() % 6;
      } else {
        monthIdx = tDate.getMonth();
      }
      if (monthIdx >= 0 && monthIdx < months.length) {
        sectorMap[secName][monthIdx] += t.quantity;
      }
    });

    requests.forEach(r => {
      if (r.status !== 'ENTREGUE') return;
      const rDate = new Date(r.deliveredAt || r.date);
      if (rDate < startDate || rDate > endDate) return;

      const secName = (r.sector || 'OUTROS').toUpperCase();
      if (!sectorMap[secName]) {
        sectorMap[secName] = new Array(months.length).fill(0);
      }

      let monthIdx = 0;
      if (months.length === 6) {
        monthIdx = rDate.getMonth() % 6;
      } else {
        monthIdx = rDate.getMonth();
      }

      const rItems = allRequestItems.filter(ri => {
        if (ri.request_id !== r.id) return false;
        if (quantitativoCategory === 'Todos') return true;
        const item = items.find(i => i.id === ri.product_id);
        return checkCategoryMatch(item?.category, quantitativoCategory);
      });

      const totalQty = rItems.reduce((acc, curr) => acc + (curr.quantity_approved || curr.quantity_requested || 0), 0);
      if (monthIdx >= 0 && monthIdx < months.length) {
        sectorMap[secName][monthIdx] += totalQty;
      }
    });

    const sectors = Object.keys(sectorMap)
      .map(name => {
        const values = sectorMap[name];
        const total = values.reduce((a, b) => a + b, 0);
        return { name, values, total };
      })
      .filter(s => quantitativoSource === 'system' ? true : s.total > 0)
      .sort((a, b) => b.total - a.total);

    const activeSectors = sectors.filter(s => s.total > 0);
    const finalSectors = activeSectors.length > 0 
      ? activeSectors 
      : (quantitativoSource === 'system' 
        ? SECTORS.slice(0, 6).map(sec => ({ name: sec.toUpperCase(), values: new Array(months.length).fill(0), total: 0 }))
        : [
          { name: 'PÃ‰ DIABÃ‰TICO', values: [384, 476, 563, 548, 572, 552], total: 3095 },
          { name: 'ILHA', values: [316, 178, 266, 310, 579, 200], total: 1849 },
          { name: 'IMAGEM', values: [351, 354, 131, 267, 505, 106], total: 1714 }
        ]);

    const activeSectorsForAnalysis = finalSectors.filter(s => s.total > 0);

    let periodText = 'no perÃ­odo analisado';
    if (quantitativoPeriodPreset === '1_semestre_2026') periodText = 'no 1Âº semestre de 2026';
    else if (quantitativoPeriodPreset === '2_semestre_2026') periodText = 'no 2Âº semestre de 2026';
    else if (quantitativoPeriodPreset === 'ano_2026') periodText = 'no ano de 2026 (total)';

    const catLabel = quantitativoCategory === 'Todos' ? 'materiais e insumos em geral' : `materiais da categoria ${quantitativoCategory.toUpperCase()}`;

    let autoAnalysis = '';
    if (activeSectorsForAnalysis.length > 0) {
      const top1 = activeSectorsForAnalysis[0];
      const top2 = activeSectorsForAnalysis[1];
      const grandTotal = activeSectorsForAnalysis.reduce((acc, s) => acc + s.total, 0);

      const monthTotals = months.map((_, idx) => activeSectorsForAnalysis.reduce((sum, sec) => sum + (sec.values[idx] || 0), 0));
      const maxMonthIdx = monthTotals.indexOf(Math.max(...monthTotals));
      const maxMonthName = months[maxMonthIdx] || 'mÃªs de pico';

      let sector2Text = '';
      if (top2 && top2.total > 0) {
        sector2Text = ` Em SEGUNDO LUGAR, destaca-se o setor de ${top2.name}, acumulando ${top2.total.toLocaleString('pt-BR')} unidades (${((top2.total / grandTotal) * 100).toFixed(1)}% do total).`;
      }

      autoAnalysis = `Verificou-se que, ${periodText}, o volume total de dispensaÃ§Ã£o para ${catLabel} foi de ${grandTotal.toLocaleString('pt-BR')} unidades. O setor com MAIOR DEMANDA foi o de ${top1.name}, apresentando ${top1.total.toLocaleString('pt-BR')} unidades dispensadas (${((top1.total / grandTotal) * 100).toFixed(1)}% do consumo total).${sector2Text} Observou-se o maior pico de dispensaÃ§Ãµes no mÃªs de ${maxMonthName}. Os dados registrados pelo sistema indicam maior concentraÃ§Ã£o assistencial nesses setores e auxiliam no planejamento das compras e estoques do almoxarifado.`;
    } else {
      autoAnalysis = `Verificou-se que, ${periodText}, nÃ£o foram registradas movimentaÃ§Ãµes de saÃ­da ou solicitaÃ§Ãµes entregues para ${catLabel} no sistema. Os controles de estoque do almoxarifado permanecem monitorando o fluxo de demandas.`;
    }

    return {
      months,
      monthColors,
      sectors: finalSectors,
      title: quantitativoTitle || (quantitativoCategory === 'Todos' ? 'QUANTITATIVO GERAL DE MATERIAIS DISPENSADOS PARA OS SETORES DO CEO' : `QUANTITATIVO DE ${quantitativoCategory.toUpperCase()} DISPENSADOS PARA OS SETORES DO CEO`),
      criticalAnalysis: quantitativoCriticalAnalysis.trim() !== '' ? quantitativoCriticalAnalysis : autoAnalysis
    };
  }, [quantitativoSource, quantitativoPeriodPreset, quantitativoCustomStart, quantitativoCustomEnd, quantitativoCategory, quantitativoTitle, quantitativoCriticalAnalysis, transactions, requests, allRequestItems, items]);

  const handleExportQuantitativoPDF = async () => {
    if (!quantitativoReportRef.current) return;
    try {
      showToast("Gerando PDF oficial do relatÃ³rio...", "info");
      const element = quantitativoReportRef.current;
      const canvas = await html2canvas(element, {
        scale: 2,
        useCORS: true,
        logging: false,
        backgroundColor: '#ffffff',
        onclone: (clonedDoc) => {
          // Hide all UI buttons, tooltips, and edit controls in the cloned document
          const pdfHideElements = clonedDoc.querySelectorAll('[data-pdf-hide="true"], button');
          pdfHideElements.forEach((el) => {
            (el as HTMLElement).style.display = 'none';
          });

          // Accurate OKLCH to RGB converter for html2canvas compatibility
          const oklchToRgb = (oklchStr: string): string => {
            try {
              const match = oklchStr.match(/oklch\(\s*([\d.%]+)\s+([\d.]+)\s+([\d.]+)(?:\s*\/\s*([\d.%]+))?\s*\)/i);
              if (!match) return '#ffffff';

              let L = parseFloat(match[1]);
              if (match[1].endsWith('%')) L /= 100;
              const C = parseFloat(match[2]);
              const H = parseFloat(match[3]);
              let A = match[4] ? parseFloat(match[4]) : 1;
              if (match[4] && match[4].endsWith('%')) A /= 100;

              const hRad = (H * Math.PI) / 180;
              const a = C * Math.cos(hRad);
              const b = C * Math.sin(hRad);

              const l_ = L + 0.3963377774 * a + 0.2158037573 * b;
              const m_ = L - 0.1055613458 * a - 0.0638541728 * b;
              const s_ = L - 0.0894841775 * a - 1.2914855480 * b;

              const l = l_ ** 3;
              const m = m_ ** 3;
              const s = s_ ** 3;

              let r = +4.0767416621 * l - 3.3077115913 * m + 0.2309699292 * s;
              let g = -1.2684380046 * l + 2.6097574011 * m - 0.3413193965 * s;
              let blue = -0.0041960863 * l - 0.7034186147 * m + 1.7076147010 * s;

              const gamma = (x: number) => (x <= 0.0031308 ? 12.92 * x : 1.055 * (Math.max(0, x) ** (1 / 2.4)) - 0.055);
              r = Math.min(255, Math.max(0, Math.round(gamma(r) * 255)));
              g = Math.min(255, Math.max(0, Math.round(gamma(g) * 255)));
              blue = Math.min(255, Math.max(0, Math.round(gamma(blue) * 255)));

              if (A < 1) {
                return `rgba(${r}, ${g}, ${blue}, ${A})`;
              }
              return `rgb(${r}, ${g}, ${blue})`;
            } catch {
              return '#ffffff';
            }
          };

          const fixStylesString = (str: string) => {
            return str
              .replace(/oklch\([^)]+\)/gi, (match) => oklchToRgb(match))
              .replace(/color-mix\([^)]+\)/gi, 'rgba(226, 232, 240, 0.8)');
          };

          // Convert oklch in <style> tags to valid rgb(...) colors so html2canvas doesn't fail or corrupt CSS variables
          const styleElements = clonedDoc.querySelectorAll('style');
          styleElements.forEach((style) => {
            if (style.textContent && (style.textContent.includes('oklch') || style.textContent.includes('color-mix'))) {
              style.textContent = fixStylesString(style.textContent);
            }
          });

          // Convert inline style attributes in cloned elements
          const allElements = clonedDoc.querySelectorAll('*');
          allElements.forEach((el) => {
            const htmlEl = el as HTMLElement;
            const styleAttr = htmlEl.getAttribute('style');
            if (styleAttr && (styleAttr.includes('oklch') || styleAttr.includes('color-mix'))) {
              htmlEl.setAttribute('style', fixStylesString(styleAttr));
            }
          });
        }
      });

      const imgData = canvas.toDataURL('image/png');
      const pdf = new jsPDF({
        orientation: 'landscape',
        unit: 'mm',
        format: 'a4'
      });
      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pdfHeight = pdf.internal.pageSize.getHeight();
      
      const margin = 10;
      const imgWidth = pdfWidth - margin * 2;
      const imgHeight = (canvas.height * imgWidth) / canvas.width;
      
      pdf.addImage(imgData, 'PNG', margin, margin, imgWidth, Math.min(imgHeight, pdfHeight - margin * 2));
      pdf.save(`Quantitativo_Insumos_Setores_${format(new Date(), 'yyyyMMdd_HHmm')}.pdf`);
      showToast("PDF oficial gerado e baixado com sucesso!", "success");
    } catch (err) {
      console.error("Erro ao gerar PDF:", err);
      showToast("Erro ao gerar PDF. Tente usar a funÃ§Ã£o de impressÃ£o.", "error");
    }
  };

  const handleExportQuantitativoExcel = () => {
    const dataToExport = quantitativoReportData.sectors.map(s => {
      const row: Record<string, any> = { 'Setor': s.name };
      quantitativoReportData.months.forEach((m, idx) => {
        row[m] = s.values[idx] || 0;
      });
      row['Total Geral'] = s.total;
      return row;
    });

    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Quantitativo por Setor");
    XLSX.writeFile(wb, `Quantitativo_Setores_${format(new Date(), 'yyyyMMdd')}.xlsx`);
    showToast("Planilha Excel exportada com sucesso!", "success");
  };

  const handleExportRoomInventoryPDF = (roomFilter: string, displayRoomName: string, filteredCategories: string[]) => {
    try {
      // @ts-ignore - jsPDF types might not be perfectly aligned with imports
      const doc = new jsPDF();
      const pageWidth = doc.internal.pageSize.width;
      
      // Header
      doc.setDrawColor(37, 99, 235); // blue-600
      doc.setLineWidth(1.5);
      doc.line(14, 15, 24, 15);
      doc.line(19, 10, 19, 20);
      
      doc.setFontSize(16);
      doc.setTextColor(28, 25, 23);
      doc.setFont('helvetica', 'bold');
      doc.text('CEO', 28, 17);
      
      doc.setFontSize(8);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(120, 113, 108);
      doc.text('CONTROLE DE ESTOQUE POR SALA', 28, 22);

      doc.setDrawColor(231, 229, 228);
      doc.setLineWidth(0.5);
      doc.line(14, 28, pageWidth - 14, 28);
      
      doc.setFontSize(14);
      doc.setTextColor(28, 25, 23);
      doc.setFont('helvetica', 'bold');
      doc.text(`Mapa de Estoque - ${displayRoomName}`, 14, 40);
      
      doc.setFontSize(9);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(120, 113, 108);
      doc.text(`Local FÃ­sico Origem: ${roomFilter}`, 14, 46);
      doc.text(`Emitido em: ${format(new Date(), 'dd/MM/yyyy HH:mm')}`, 14, 52);

      // Filter items by room and categories
      const roomItems = items.filter(i => {
        // Ignorar excluÃ­dos ou sem estoque
        if (i.deletedAt || i.quantity <= 0) return false;
        
        // NormalizaÃ§Ã£o para comparaÃ§Ã£o robusta
        const itemRoom = (i.room || 'Almoxarifado Principal').trim().toLowerCase();
        const targetRoom = roomFilter.trim().toLowerCase();
        
        const matchesRoom = itemRoom === targetRoom;
        
        // Se nenhuma categoria selecionada, mostra tudo da sala. Se selecionadas, filtra.
        const matchesCategory = filteredCategories.length === 0 || 
                               (i.category && filteredCategories.some(cat => 
                                 cat.trim().toLowerCase() === i.category?.trim().toLowerCase()
                               ));
        
        return matchesRoom && matchesCategory;
      }).sort((a, b) => a.name.localeCompare(b.name));

      if (roomItems.length === 0) {
        doc.setFontSize(10);
        doc.setTextColor(150, 150, 150);
        doc.text('NENHUM ITEM ENCONTRADO PARA OS FILTROS SELECIONADOS.', 14, 70);
      } else {
        const tableData = roomItems.map(item => {
          const daysToExpiry = item.expiry_date && item.expiry_date !== 'Indeterminada' 
            ? differenceInDays(new Date(item.expiry_date), new Date()) 
            : null;
            
          let expiryStatus = '-';
          if (daysToExpiry !== null) {
            if (daysToExpiry < 0) expiryStatus = 'VENCIDO';
            else if (daysToExpiry <= 30) expiryStatus = 'CRÃTICO';
            else expiryStatus = `${daysToExpiry} dias`;
          } else if (item.expiry_date === 'Indeterminada') {
            expiryStatus = 'Indeterminada';
          }

          return [
            item.name,
            item.batch_number || '-',
            item.category || '-',
            { content: item.quantity.toString(), styles: { fontStyle: 'bold' as any, halign: 'center' as any } },
            item.expiry_date || '-',
            { content: expiryStatus, styles: { halign: 'center' as any } }
          ];
        });

        autoTable(doc, {
          startY: 60,
          head: [['Produto', 'Lote', 'Categoria', 'Estoque', 'Validade', 'Status (Dias)']],
          body: tableData,
          theme: 'striped',
          headStyles: { 
            fillColor: [28, 25, 23],
            textColor: [255, 255, 255],
            fontSize: 8,
            fontStyle: 'bold',
            halign: 'center'
          },
          styles: { fontSize: 8, cellPadding: 2.5 },
          columnStyles: {
            0: { cellWidth: 'auto' },
            1: { cellWidth: 25 },
            2: { cellWidth: 35 },
            3: { cellWidth: 20 },
            4: { cellWidth: 25 },
            5: { cellWidth: 30 }
          },
          margin: { horizontal: 14 }
        });
      }
      
      const safeRoomName = displayRoomName.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, '-');
      doc.save(`mapa-sala-${safeRoomName}-${format(new Date(), 'yyyy-MM-dd')}.pdf`);
      showToast("Documento de porta gerado com sucesso!", "success");
    } catch (error) {
      console.error("PDF Error:", error);
      showToast("Erro ao gerar PDF", "error");
    }
  };

  const getImageDataURL = async (url: string): Promise<string> => {
    try {
      const baseUrl = typeof window !== 'undefined' ? window.location.origin : '';
      const fullUrl = url.startsWith('http') ? url : `${baseUrl}${url.startsWith('/') ? '' : '/'}${url}`;
      
      const response = await fetch(fullUrl);
      if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
      const blob = await response.blob();
      
      if (blob.size < 500) {
        throw new Error(`Imagem muito pequena: ${blob.size} bytes`);
      }

      // Converte para JPEG via Canvas para evitar erros de signature no jsPDF
      return new Promise((resolve, reject) => {
        const img = new Image();
        img.crossOrigin = 'anonymous';
        
        const timeout = setTimeout(() => {
          reject(new Error("Timeout carregando imagem"));
        }, 8000);

        img.onload = () => {
          clearTimeout(timeout);
          try {
            const canvas = document.createElement('canvas');
            canvas.width = img.width;
            canvas.height = img.height;
            const ctx = canvas.getContext('2d');
            if (!ctx) {
              reject(new Error("Erro ao criar contexto de canvas"));
              return;
            }
            // Fundo branco para imagens transparentes
            ctx.fillStyle = "#FFFFFF";
            ctx.fillRect(0, 0, canvas.width, canvas.height);
            ctx.drawImage(img, 0, 0);
            
            // ForÃ§amos o formato JPEG com qualidade alta
            const dataUrl = canvas.toDataURL('image/jpeg', 0.95);
            resolve(dataUrl);
            URL.revokeObjectURL(img.src);
          } catch (e) {
            reject(e);
          }
        };

        img.onerror = () => {
          clearTimeout(timeout);
          // Se o Canvas falhar, tenta FileReader direto como Ãºltimo recurso
          const reader = new FileReader();
          reader.onloadend = () => resolve(reader.result as string);
          reader.onerror = () => reject(new Error("Erro ao processar imagem"));
          reader.readAsDataURL(blob);
          URL.revokeObjectURL(img.src);
        };

        img.src = URL.createObjectURL(blob);
      });
    } catch (err) {
      console.error(`[PDF] Erro em getImageDataURL (${url}):`, err);
      throw err;
    }
  };

  const handleExportDonationTermPDF = async (data: {
    donatingUnitName?: string;
    receivingUnit: { name: string; address: string; cnpj: string };
    items: { product_name: string; quantity: number; batch_number?: string }[];
    revisionDate: string;
    donationNumber?: string;
    date: string;
  }) => {
    try {
      showToast("Gerando Termo de DoaÃ§Ã£o...", "info");

      let base64Image = letterheadImage || "";
      
      // Se nÃ£o houver imagem personalizada, tenta carregar a padrÃ£o
      if (!base64Image) {
        try {
          base64Image = await getImageDataURL("/official_letterhead.svg");
        } catch (err) {
          console.warn("Could not load logo image for Donation Term, using fallback text header:", err);
        }
      }
      
      // @ts-ignore
      const doc = new jsPDF();
      const pageWidth = doc.internal.pageSize.width;
      const pageHeight = doc.internal.pageSize.height;
      const margin = 20;

      const drawLetterhead = (pdfDoc: any) => {
        if (base64Image) {
          try {
            console.log("[PDF] Desenhando imagem de papel timbrado no Termo de DoaÃ§Ã£o");
            const format = base64Image.includes('image/png') ? 'PNG' : 'JPEG';
            pdfDoc.addImage(base64Image, format, 0, 0, pageWidth, pageHeight, undefined, 'FAST');
            return;
          } catch (e) {
            console.error("Error adding letterhead image to Donation Term:", e);
          }
        }
        
        console.log("[PDF] Usando cabeÃ§alho padrÃ£o com 3 logos retangulares expandidos no Termo de DoaÃ§Ã£o");
        const docLogo = appRectangularLogo || appLogo;
        const logoWidth = 50;
        const logoHeight = 16;
        const logoY = 10;
        
        // 1. LOGO ALMOXARIFADO (Left - Rectangular)
        if (docLogo) {
          try {
            const format = docLogo.includes('image/png') ? 'PNG' : 'JPEG';
            pdfDoc.addImage(docLogo, format, margin, logoY, logoWidth, logoHeight, undefined, 'FAST');
          } catch (e) {
            console.error("Error adding docLogo to Donation Term:", e);
          }
        } else {
          pdfDoc.setFillColor(240, 253, 244);
          pdfDoc.roundedRect(margin, logoY, logoWidth, logoHeight, 2, 2, 'F');
          pdfDoc.setFontSize(8);
          pdfDoc.setFont('helvetica', 'bold');
          pdfDoc.setTextColor(22, 101, 52);
          pdfDoc.text('ALMOXARIFADO', margin + (logoWidth / 2), logoY + 10, { align: 'center' });
        }

        // 2. LOGO POLICLÃNICA (Center - Rectangular)
        const centerX = (pageWidth / 2) - (logoWidth / 2);
        if (policlinicaLogo) {
          try {
            const format = policlinicaLogo.includes('image/png') ? 'PNG' : 'JPEG';
            pdfDoc.addImage(policlinicaLogo, format, centerX, logoY, logoWidth, logoHeight, undefined, 'FAST');
          } catch (e) {
            console.error("Error adding policlinicaLogo to Donation Term:", e);
          }
        } else {
          pdfDoc.setFillColor(240, 249, 255);
          pdfDoc.roundedRect(centerX, logoY, logoWidth, logoHeight, 2, 2, 'F');
          pdfDoc.setFontSize(8);
          pdfDoc.setFont('helvetica', 'bold');
          pdfDoc.setTextColor(3, 105, 161);
          pdfDoc.text('CEO - CENTRO DE ESPECIALIDADES ODONTOLÃ“GICAS', centerX + (logoWidth / 2), logoY + 10, { align: 'center' });
        }

        // 3. LOGO CONSÃ“RCIO CPSMS (Right - Rectangular)
        const consorcioWidth = 56;
        const consorcioHeight = 18;
        const consorcioY = 9;
        const rightX = pageWidth - margin - consorcioWidth;
        if (consorcioLogo) {
          try {
            const format = consorcioLogo.includes('image/png') ? 'PNG' : 'JPEG';
            pdfDoc.addImage(consorcioLogo, format, rightX, consorcioY, consorcioWidth, consorcioHeight, undefined, 'FAST');
          } catch (e) {
            console.error("Error adding consorcioLogo to Donation Term:", e);
          }
        } else {
          pdfDoc.setFillColor(255, 247, 237);
          pdfDoc.roundedRect(rightX, consorcioY, consorcioWidth, consorcioHeight, 2, 2, 'F');
          pdfDoc.setFontSize(8);
          pdfDoc.setFont('helvetica', 'bold');
          pdfDoc.setTextColor(194, 65, 12);
          pdfDoc.text('CONSÃ“RCIO CPSMS', rightX + (consorcioWidth / 2), consorcioY + 11, { align: 'center' });
        }

        pdfDoc.setDrawColor(226, 232, 240);
        pdfDoc.setLineWidth(0.5);
        pdfDoc.line(margin, 29, pageWidth - margin, 29);

        // Footer
        pdfDoc.setFontSize(7.5);
        pdfDoc.setTextColor(120, 113, 108);
        pdfDoc.setFont('helvetica', 'normal');
        const footer1 = 'CEO - Centro de Especialidades OdontolÃ³gicas.';
        const footer2 = 'Fone: (88) 3614-3156 | Fax: (88) 3614-3245 | cpsms.ce.gov.br';
        pdfDoc.text(footer1, pageWidth / 2, pageHeight - 12, { align: 'center' });
        pdfDoc.text(footer2, pageWidth / 2, pageHeight - 8, { align: 'center' });
      };

      const formatTitleCase = (str: string) => {
        if (!str) return '';
        const lower = str.toLowerCase();
        const minorWords = ['de', 'da', 'do', 'das', 'dos', 'e', 'em', 'para'];
        return lower.split(' ').map((word, index) => {
          if (index > 0 && minorWords.includes(word)) return word;
          return word.charAt(0).toUpperCase() + word.slice(1);
        }).join(' ');
      };

      const donorName = formatTitleCase(data.donatingUnitName || 'CEO - Centro de Especialidades OdontolÃ³gicas');
      const receivingName = formatTitleCase(data.receivingUnit.name);
      const receivingAddress = data.receivingUnit.address;
      const receivingCNPJ = data.receivingUnit.cnpj;

      drawLetterhead(doc);

      // --- TITLE & DATA DE EMISSÃƒO BELOW LOGOS ---
      doc.setFontSize(13);
      doc.setTextColor(17, 24, 39);
      doc.setFont('helvetica', 'bold');
      doc.text('TERMO DE DOAÃ‡ÃƒO DE MATERIAIS E INSUMOS', pageWidth / 2, 35, { align: 'center' });
      
      doc.setFontSize(8.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(100, 116, 139);
      doc.text(`Data de EmissÃ£o: ${format(new Date(), 'dd/MM/yyyy HH:mm')}`, pageWidth / 2, 40, { align: 'center' });

      // --- DOCUMENT METADATA RIGHT-ALIGNED ---
      doc.setFontSize(8);
      doc.setTextColor(107, 114, 128);
      doc.setFont('helvetica', 'normal');
      doc.text('CÃ³digo: TERMO-ALMOX', pageWidth - margin, 46, { align: 'right' });
      doc.text(`Data de ImplantaÃ§Ã£o: ${format(new Date(), 'dd/MM/yyyy')}`, pageWidth - margin, 50, { align: 'right' });
      doc.text(`Ãšltima RevisÃ£o: ${data.revisionDate || '---'}`, pageWidth - margin, 54, { align: 'right' });
      
      if (data.donationNumber) {
        doc.setFontSize(9);
        doc.setTextColor(31, 41, 55);
        doc.setFont('helvetica', 'bold');
        doc.text(`Termo nÂº: ${data.donationNumber}`, pageWidth - margin, 59, { align: 'right' });
      }

      doc.setDrawColor(209, 213, 219);
      doc.setLineWidth(0.2);
      doc.line(margin, 63, pageWidth - margin, 63);

      // --- CONTENT ---
      doc.setFontSize(10);
      doc.setTextColor(31, 41, 55);
      doc.setFont('helvetica', 'normal');
      
      const donationText = `A ${donorName}, inscrita sob o CNPJ nÂº 12.208.466/0001-66, por intermÃ©dio de seu Setor de Almoxarifado, formaliza por este instrumento a doaÃ§Ã£o Ã  unidade ${receivingName}, situada em ${receivingAddress}, inscrita sob o CNPJ nÂº ${receivingCNPJ}, dos materiais e insumos abaixo discriminados. A presente cessÃ£o justifica-se pela otimizaÃ§Ã£o de estoque em virtude da reduÃ§Ã£o de demanda interna e proximidade do prazo de validade, assegurando a destinaÃ§Ã£o Ãºtil dos itens.`;
      
      const textWidth = pageWidth - (margin * 2);
      const textLines = doc.splitTextToSize(donationText, textWidth);
      doc.text(textLines, margin, 72, { 
        align: 'justify', 
        maxWidth: textWidth,
        lineHeightFactor: 1.5 
      });

      const tableStartY = 85 + (textLines.length * 7) + 5;

      autoTable(doc, {
        startY: tableStartY,
        margin: { left: margin, right: margin },
        head: [['DescriÃ§Ã£o do Material', 'Lote', 'Qtd Doada', 'ConferÃªncia']],
        body: data.items.map(i => [
          i.product_name, 
          (i.batch_number || '---').toUpperCase(), 
          i.quantity.toString(), 
          ' '
        ]),
        theme: 'grid',
        headStyles: { 
          fillColor: [243, 244, 246], 
          textColor: [31, 41, 55],
          fontStyle: 'bold',
          halign: 'left',
          fontSize: 9,
          lineWidth: 0.1,
          lineColor: [209, 213, 219]
        },
        styles: { 
          fontSize: 8, 
          cellPadding: 4,
          lineColor: [209, 213, 219],
          lineWidth: 0.1,
          textColor: [55, 65, 81]
        },
        columnStyles: {
          0: { cellWidth: 'auto' },
          1: { cellWidth: 35, halign: 'center', fontStyle: 'bold' },
          2: { cellWidth: 25, halign: 'center', fontStyle: 'bold' },
          3: { cellWidth: 35, halign: 'center' }
        },
        didDrawPage: (data) => {
          if (data.pageNumber > 1) {
            drawLetterhead(doc);
          }
        }
      });

      const tableFinalY = (doc as any).lastAutoTable.finalY;
      let signAreaY = tableFinalY + 15;

      if (signAreaY + 50 > pageHeight - 20) {
        doc.addPage();
        signAreaY = 40;
      }

      doc.setFontSize(10);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(31, 41, 55);
      const formattedDate = format(new Date(data.date), "dd 'de' MMMM 'de' yyyy", { locale: ptBR });
      doc.text(`Sobral-CE, ${formattedDate}.`, pageWidth / 2, signAreaY, { align: 'center' });

      const signY = signAreaY + 25;
      doc.setDrawColor(156, 163, 175);
      doc.setLineWidth(0.5);
      const signLineW = 75;
      
      doc.line(margin, signY, margin + signLineW, signY);
      doc.setFontSize(8);
      doc.setFont('helvetica', 'bold');
      doc.text(donorName, margin + (signLineW / 2), signY + 5, { align: 'center' });
      doc.setFont('helvetica', 'normal');
      doc.text('Unidade Doadora', margin + (signLineW / 2), signY + 10, { align: 'center' });
      doc.text('(assinatura e carimbo)', margin + (signLineW / 2), signY + 14, { align: 'center' });
      
      doc.line(pageWidth - margin - signLineW, signY, pageWidth - margin, signY);
      doc.setFontSize(8);
      doc.setFont('helvetica', 'bold');
      doc.text(receivingName, pageWidth - margin - (signLineW / 2), signY + 5, { align: 'center' });
      doc.setFont('helvetica', 'normal');
      doc.text('Unidade Receptora', pageWidth - margin - (signLineW / 2), signY + 10, { align: 'center' });
      doc.text('(assinatura e carimbo)', pageWidth - margin - (signLineW / 2), signY + 14, { align: 'center' });

      doc.save(`Termo_Doacao_${data.receivingUnit.name.replace(/\s+/g, '_')}_${format(new Date(), 'dd-MM-yyyy')}.pdf`);
      showToast("Termo de DoaÃ§Ã£o gerado com sucesso!", "success");
    } catch (error) {
      console.error('Erro ao exportar PDF de DoaÃ§Ã£o:', error);
      alert('Ocorreu um erro ao gerar o Termo de DoaÃ§Ã£o.');
    }
  };

  const drawPDFLetterhead = (doc: any, title?: string, subtitle?: string): number => {
    const pageWidth = doc.internal.pageSize.width;
    let startY = 14;

    if (letterheadImage) {
      try {
        const format = letterheadImage.includes('image/png') ? 'PNG' : 'JPEG';
        doc.addImage(letterheadImage, format, 14, 8, pageWidth - 28, 28, undefined, 'FAST');
        startY = 40;
      } catch (e) {
        console.warn("Could not render letterheadImage on PDF report:", e);
      }
    } else {
      // Perfectly aligned 3-logo horizontal header row (homogeneous rectangular logos)
      const docLogo = appRectangularLogo || appLogo;
      const logoWidth = 50;
      const logoHeight = 16;
      const logoY = 10;
      
      // 1. LEFT LOGO: Logo Almoxarifado / Sistema (Rectangular)
      if (docLogo) {
        try {
          const format = docLogo.includes('image/png') ? 'PNG' : 'JPEG';
          doc.addImage(docLogo, format, 14, logoY, logoWidth, logoHeight, undefined, 'FAST');
        } catch (e) {
          console.warn("Could not render logo on PDF report:", e);
        }
      } else {
        doc.setFillColor(240, 253, 244);
        doc.roundedRect(14, logoY, logoWidth, logoHeight, 2, 2, 'F');
        doc.setFontSize(8);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(22, 101, 52);
        doc.text('ALMOXARIFADO', 14 + (logoWidth / 2), logoY + 10, { align: 'center' });
      }

      // 2. CENTER LOGO: Logo da PoliclÃ­nica (Rectangular)
      const centerX = (pageWidth / 2) - (logoWidth / 2);
      if (policlinicaLogo) {
        try {
          const format = policlinicaLogo.includes('image/png') ? 'PNG' : 'JPEG';
          doc.addImage(policlinicaLogo, format, centerX, logoY, logoWidth, logoHeight, undefined, 'FAST');
        } catch (e) {
          console.warn("Could not render policlinicaLogo on PDF:", e);
        }
      } else {
        doc.setFillColor(240, 249, 255);
        doc.roundedRect(centerX, logoY, logoWidth, logoHeight, 2, 2, 'F');
        doc.setFontSize(8);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(3, 105, 161);
        doc.text('CEO - CENTRO DE ESPECIALIDADES ODONTOLÃ“GICAS', centerX + (logoWidth / 2), logoY + 10, { align: 'center' });
      }

      // 3. RIGHT LOGO: Logo do ConsÃ³rcio CPSMS (Rectangular)
      const consorcioWidth = 56;
      const consorcioHeight = 18;
      const consorcioY = 9;
      const rightX = pageWidth - 14 - consorcioWidth;
      if (consorcioLogo) {
        try {
          const format = consorcioLogo.includes('image/png') ? 'PNG' : 'JPEG';
          doc.addImage(consorcioLogo, format, rightX, consorcioY, consorcioWidth, consorcioHeight, undefined, 'FAST');
        } catch (e) {
          console.warn("Could not render consorcioLogo on PDF:", e);
        }
      } else {
        doc.setFillColor(255, 247, 237);
        doc.roundedRect(rightX, consorcioY, consorcioWidth, consorcioHeight, 2, 2, 'F');
        doc.setFontSize(8);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(194, 65, 12);
        doc.text('CONSÃ“RCIO CPSMS', rightX + (consorcioWidth / 2), consorcioY + 11, { align: 'center' });
      }

      // Divider Line
      doc.setDrawColor(226, 232, 240);
      doc.setLineWidth(0.5);
      doc.line(14, 29, pageWidth - 14, 29);

      startY = 35;
    }

    if (title) {
      doc.setFontSize(13);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(28, 25, 23);
      doc.text(title, pageWidth / 2, startY, { align: 'center' });
      startY += 5;
    }

    doc.setFontSize(8.5);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(100, 116, 139);
    doc.text(`Data de EmissÃ£o: ${format(new Date(), 'dd/MM/yyyy HH:mm')}`, pageWidth / 2, startY, { align: 'center' });
    startY += 7;

    if (subtitle) {
      doc.setFontSize(8.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(120, 113, 108);
      doc.text(subtitle, pageWidth / 2, startY, { align: 'center' });
      startY += 7;
    }

    return startY;
  };

  const [appLogo, setAppLogo] = useState<string | null>(null);
  const [appRectangularLogo, setAppRectangularLogo] = useState<string | null>(null);
  const [policlinicaLogo, setPoliclinicaLogo] = useState<string | null>(null);
  const [consorcioLogo, setConsorcioLogo] = useState<string | null>(null);

  // Load app logo & letterhead from localStorage & Firestore on mount
  useEffect(() => {
    const savedLogo = localStorage.getItem('app_logo_base64');
    if (savedLogo) setAppLogo(savedLogo);

    const savedRectLogo = localStorage.getItem('app_rectangular_logo_base64');
    if (savedRectLogo) setAppRectangularLogo(savedRectLogo);

    const savedPoliLogo = localStorage.getItem('policlinica_logo_base64');
    if (savedPoliLogo) setPoliclinicaLogo(savedPoliLogo);

    const savedConsLogo = localStorage.getItem('consorcio_logo_base64');
    if (savedConsLogo) setConsorcioLogo(savedConsLogo);

    const savedLetterhead = localStorage.getItem('letterhead_image_base64');
    if (savedLetterhead) setLetterheadImage(savedLetterhead);

    const unsubscribe = onSnapshot(doc(db, 'settings', 'general'), (snapshot) => {
      if (snapshot.exists()) {
        const data = snapshot.data();
        if (data.appLogo) {
          setAppLogo(data.appLogo);
          localStorage.setItem('app_logo_base64', data.appLogo);
        } else {
          setAppLogo(null);
          localStorage.removeItem('app_logo_base64');
        }

        if (data.appRectangularLogo) {
          setAppRectangularLogo(data.appRectangularLogo);
          localStorage.setItem('app_rectangular_logo_base64', data.appRectangularLogo);
        } else {
          setAppRectangularLogo(null);
          localStorage.removeItem('app_rectangular_logo_base64');
        }

        if (data.policlinicaLogo) {
          setPoliclinicaLogo(data.policlinicaLogo);
          localStorage.setItem('policlinica_logo_base64', data.policlinicaLogo);
        } else {
          setPoliclinicaLogo(null);
          localStorage.removeItem('policlinica_logo_base64');
        }

        if (data.consorcioLogo) {
          setConsorcioLogo(data.consorcioLogo);
          localStorage.setItem('consorcio_logo_base64', data.consorcioLogo);
        } else {
          setConsorcioLogo(null);
          localStorage.removeItem('consorcio_logo_base64');
        }

        if (data.letterheadImage) {
          setLetterheadImage(data.letterheadImage);
          localStorage.setItem('letterhead_image_base64', data.letterheadImage);
        } else {
          setLetterheadImage(null);
          localStorage.removeItem('letterhead_image_base64');
        }
      }
    }, (err) => {
      console.warn("Could not listen to settings/general:", err);
    });

    return () => unsubscribe();
  }, []);

  const handleLetterheadUpload = async (file: File) => {
    if (file.size > 5 * 1024 * 1024) {
      showToast("Imagem muito grande. MÃ¡ximo 5MB.", "error");
      return;
    }
    const reader = new FileReader();
    reader.onloadend = async () => {
      const base64 = reader.result as string;
      setLetterheadImage(base64);
      localStorage.setItem('letterhead_image_base64', base64);
      try {
        await setDoc(doc(db, 'settings', 'general'), { letterheadImage: base64 }, { merge: true });
        showToast("Papel timbrado atualizado e salvo com sucesso!", "success");
      } catch (err) {
        console.error("Erro ao salvar papel timbrado no Firestore:", err);
        showToast("Papel timbrado atualizado!", "success");
      }
    };
    reader.readAsDataURL(file);
  };

  const handleRemoveLetterhead = async () => {
    setLetterheadImage(null);
    localStorage.removeItem('letterhead_image_base64');
    try {
      await setDoc(doc(db, 'settings', 'general'), { letterheadImage: deleteField() }, { merge: true });
      showToast("Papel timbrado removido com sucesso!", "success");
    } catch (err) {
      console.error("Erro ao remover papel timbrado do Firestore:", err);
      showToast("Papel timbrado removido!", "success");
    }
  };

  const handleLogoUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.size > 2 * 1024 * 1024) {
        showToast("Imagem muito grande. MÃ¡ximo 2MB.", "error");
        return;
      }
      const reader = new FileReader();
      reader.onloadend = async () => {
        const base64 = reader.result as string;
        setAppLogo(base64);
        localStorage.setItem('app_logo_base64', base64);
        try {
          await setDoc(doc(db, 'settings', 'general'), { appLogo: base64 }, { merge: true });
          showToast("Logo quadrada do sistema atualizada com sucesso!", "success");
        } catch (err) {
          console.error("Erro ao salvar no Firestore:", err);
          showToast("Logo atualizada!", "success");
        }
      };
      reader.readAsDataURL(file);
    }
  };

  const handleRemoveLogo = async () => {
    setAppLogo(null);
    localStorage.removeItem('app_logo_base64');
    try {
      await setDoc(doc(db, 'settings', 'general'), { appLogo: deleteField() }, { merge: true });
      showToast("Logo quadrada removida com sucesso!", "success");
    } catch (err) {
      console.error("Erro ao remover logo do Firestore:", err);
      showToast("Logo removida!", "success");
    }
  };

  const handleRectangularLogoUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.size > 2 * 1024 * 1024) {
        showToast("Imagem muito grande. MÃ¡ximo 2MB.", "error");
        return;
      }
      const reader = new FileReader();
      reader.onloadend = async () => {
        const base64 = reader.result as string;
        setAppRectangularLogo(base64);
        localStorage.setItem('app_rectangular_logo_base64', base64);
        try {
          await setDoc(doc(db, 'settings', 'general'), { appRectangularLogo: base64 }, { merge: true });
          showToast("Logo retangular atualizada com sucesso!", "success");
        } catch (err) {
          console.error("Erro ao salvar no Firestore:", err);
          showToast("Logo retangular atualizada!", "success");
        }
      };
      reader.readAsDataURL(file);
    }
  };

  const handleRemoveRectangularLogo = async () => {
    setAppRectangularLogo(null);
    localStorage.removeItem('app_rectangular_logo_base64');
    try {
      await setDoc(doc(db, 'settings', 'general'), { appRectangularLogo: deleteField() }, { merge: true });
      showToast("Logo retangular removida com sucesso!", "success");
    } catch (err) {
      console.error("Erro ao remover logo retangular do Firestore:", err);
      showToast("Logo retangular removida!", "success");
    }
  };

  const handlePoliclinicaLogoUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.size > 2 * 1024 * 1024) {
        showToast("Imagem muito grande. MÃ¡ximo 2MB.", "error");
        return;
      }
      const reader = new FileReader();
      reader.onloadend = async () => {
        const base64 = reader.result as string;
        setPoliclinicaLogo(base64);
        localStorage.setItem('policlinica_logo_base64', base64);
        try {
          await setDoc(doc(db, 'settings', 'general'), { policlinicaLogo: base64 }, { merge: true });
          showToast("Logo do CEO atualizada com sucesso!", "success");
        } catch (err) {
          console.error("Erro ao salvar no Firestore:", err);
          showToast("Logo do CEO atualizada!", "success");
        }
      };
      reader.readAsDataURL(file);
    }
  };

  const handleRemovePoliclinicaLogo = async () => {
    setPoliclinicaLogo(null);
    localStorage.removeItem('policlinica_logo_base64');
    try {
      await setDoc(doc(db, 'settings', 'general'), { policlinicaLogo: deleteField() }, { merge: true });
      showToast("Logo do CEO removida com sucesso!", "success");
    } catch (err) {
      console.error("Erro ao remover logo da PoliclÃ­nica do Firestore:", err);
      showToast("Logo do CEO removida!", "success");
    }
  };

  const handleConsorcioLogoUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.size > 2 * 1024 * 1024) {
        showToast("Imagem muito grande. MÃ¡ximo 2MB.", "error");
        return;
      }
      const reader = new FileReader();
      reader.onloadend = async () => {
        const base64 = reader.result as string;
        setConsorcioLogo(base64);
        localStorage.setItem('consorcio_logo_base64', base64);
        try {
          await setDoc(doc(db, 'settings', 'general'), { consorcioLogo: base64 }, { merge: true });
          showToast("Logo do ConsÃ³rcio CPSMS atualizada com sucesso!", "success");
        } catch (err) {
          console.error("Erro ao salvar no Firestore:", err);
          showToast("Logo do ConsÃ³rcio atualizada!", "success");
        }
      };
      reader.readAsDataURL(file);
    }
  };

  const handleRemoveConsorcioLogo = async () => {
    setConsorcioLogo(null);
    localStorage.removeItem('consorcio_logo_base64');
    try {
      await setDoc(doc(db, 'settings', 'general'), { consorcioLogo: deleteField() }, { merge: true });
      showToast("Logo do ConsÃ³rcio removida com sucesso!", "success");
    } catch (err) {
      console.error("Erro ao remover logo do ConsÃ³rcio do Firestore:", err);
      showToast("Logo do ConsÃ³rcio removida!", "success");
    }
  };

  const handleExportDeliveryReceiptPDF = async (data: {
    sector: string;
    items: { product_name: string; quantity: number; batch_number?: string }[];
    requestId?: string;
    date: string;
  }) => {
    try {
      showToast("Gerando Recibo...", "info");
      
      // @ts-ignore
      const doc = new jsPDF();
      const pageWidth = doc.internal.pageSize.width;
      const pageHeight = doc.internal.pageSize.height;
      const margin = 14;

      const drawLetterhead = (pdfDoc: any) => {
        const docLogo = appRectangularLogo || appLogo;
        const logoWidth = 50;
        const logoHeight = 16;
        const logoY = 10;
        
        // 1. LOGO ALMOXARIFADO (Left - Rectangular)
        if (docLogo) {
          try {
            const format = docLogo.includes('image/png') ? 'PNG' : 'JPEG';
            pdfDoc.addImage(docLogo, format, margin, logoY, logoWidth, logoHeight, undefined, 'FAST');
          } catch (e) {
            console.error("Error adding logo to Delivery Receipt:", e);
          }
        } else {
          pdfDoc.setFillColor(240, 253, 244);
          pdfDoc.roundedRect(margin, logoY, logoWidth, logoHeight, 2, 2, 'F');
          pdfDoc.setFontSize(8);
          pdfDoc.setFont('helvetica', 'bold');
          pdfDoc.setTextColor(22, 101, 52);
          pdfDoc.text('ALMOXARIFADO', margin + (logoWidth / 2), logoY + 10, { align: 'center' });
        }

        // 2. LOGO POLICLÃNICA (Center - Rectangular)
        const centerX = (pageWidth / 2) - (logoWidth / 2);
        if (policlinicaLogo) {
          try {
            const format = policlinicaLogo.includes('image/png') ? 'PNG' : 'JPEG';
            pdfDoc.addImage(policlinicaLogo, format, centerX, logoY, logoWidth, logoHeight, undefined, 'FAST');
          } catch (e) {
            console.error("Error adding policlinicaLogo to Delivery Receipt:", e);
          }
        } else {
          pdfDoc.setFillColor(240, 249, 255);
          pdfDoc.roundedRect(centerX, logoY, logoWidth, logoHeight, 2, 2, 'F');
          pdfDoc.setFontSize(8);
          pdfDoc.setFont('helvetica', 'bold');
          pdfDoc.setTextColor(3, 105, 161);
          pdfDoc.text('CEO - CENTRO DE ESPECIALIDADES ODONTOLÃ“GICAS', centerX + (logoWidth / 2), logoY + 10, { align: 'center' });
        }

        // 3. LOGO CONSÃ“RCIO CPSMS (Right - Rectangular)
        const consorcioWidth = 56;
        const consorcioHeight = 18;
        const consorcioY = 9;
        const rightX = pageWidth - margin - consorcioWidth;
        if (consorcioLogo) {
          try {
            const format = consorcioLogo.includes('image/png') ? 'PNG' : 'JPEG';
            pdfDoc.addImage(consorcioLogo, format, rightX, consorcioY, consorcioWidth, consorcioHeight, undefined, 'FAST');
          } catch (e) {
            console.error("Error adding consorcioLogo to Delivery Receipt:", e);
          }
        } else {
          pdfDoc.setFillColor(255, 247, 237);
          pdfDoc.roundedRect(rightX, consorcioY, consorcioWidth, consorcioHeight, 2, 2, 'F');
          pdfDoc.setFontSize(8);
          pdfDoc.setFont('helvetica', 'bold');
          pdfDoc.setTextColor(194, 65, 12);
          pdfDoc.text('CONSÃ“RCIO CPSMS', rightX + (consorcioWidth / 2), consorcioY + 11, { align: 'center' });
        }

        pdfDoc.setDrawColor(226, 232, 240);
        pdfDoc.setLineWidth(0.5);
        pdfDoc.line(margin, 29, pageWidth - margin, 29);

        // Footer
        pdfDoc.setFontSize(7.5);
        pdfDoc.setTextColor(120, 113, 108);
        pdfDoc.setFont('helvetica', 'normal');
        const footer1 = 'CEO - Centro de Especialidades OdontolÃ³gicas.';
        const footer2 = 'Fone: (88) 3614-3156 | Fax: (88) 3614-3245 | cpsms.ce.gov.br';
        pdfDoc.text(footer1, pageWidth / 2, pageHeight - 12, { align: 'center' });
        pdfDoc.text(footer2, pageWidth / 2, pageHeight - 8, { align: 'center' });
      };

      drawLetterhead(doc);
      
      // Document Title & Emission Date directly below logos
      doc.setFontSize(13);
      doc.setTextColor(28, 25, 23);
      doc.setFont('helvetica', 'bold');
      doc.text('RECIBO DE ENTREGA DE MATERIAL', pageWidth / 2, 35, { align: 'center' });
      
      doc.setFontSize(8.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(100, 116, 139);
      doc.text(`Data de EmissÃ£o: ${format(new Date(), 'dd/MM/yyyy HH:mm')}`, pageWidth / 2, 40, { align: 'center' });

      // Stylized blue separator
      doc.setDrawColor(0, 139, 190);
      doc.setLineWidth(0.5);
      doc.line(14, 44, pageWidth - 14, 44);

      // Info Card
      doc.setFillColor(248, 250, 252);
      doc.roundedRect(14, 48, pageWidth - 28, 20, 2, 2, 'F');
      
      doc.setFontSize(9);
      doc.setTextColor(71, 85, 105);
      doc.setFont('helvetica', 'bold');
      doc.text('SETOR DESTINO:', 19, 60);
      doc.text('REFERÃŠNCIA:', 19, 68);
      
      doc.setTextColor(30, 41, 59);
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(11);
      doc.text(data.sector.toUpperCase(), 52, 60);
      
      doc.setFontSize(9);
      doc.setFont('helvetica', 'normal');
      doc.text(data.requestId ? `SolicitaÃ§Ã£o #${data.requestId.slice(-5).toUpperCase()}` : 'Baixa Direta no Sistema', 52, 68);
      
      doc.text('DATA DA SAÃDA:', pageWidth - 80, 68);
      doc.setFont('helvetica', 'bold');
      doc.text(format(new Date(data.date), 'dd/MM/yyyy'), pageWidth - 50, 68);

      // Materials Table
      const tableData = data.items.map(i => [
        i.product_name.toUpperCase(), 
        (i.batch_number || '---').toUpperCase(),
        i.quantity.toString(), 
        '_________________'
      ]);
      
      autoTable(doc, {
        startY: 80,
        head: [['DESCRIÃ‡ÃƒO DO MATERIAL', 'LOTE', 'QTD ENTREGUE', 'CONFERÃŠNCIA']],
        body: tableData,
        theme: 'grid',
        headStyles: { 
          fillColor: [30, 41, 59], 
          textColor: [255, 255, 255],
          fontStyle: 'bold',
          halign: 'center',
          fontSize: 9
        },
        styles: { 
          fontSize: 8, 
          cellPadding: 4,
          lineColor: [200, 200, 200],
          lineWidth: 0.1
        },
        columnStyles: {
          0: { cellWidth: 'auto' },
          1: { cellWidth: 35, halign: 'center', fontStyle: 'bold' },
          2: { cellWidth: 30, halign: 'center', fontStyle: 'bold' },
          3: { cellWidth: 40, halign: 'center' }
        },
        alternateRowStyles: {
          fillColor: [252, 252, 252]
        },
        didDrawPage: (data) => {
          if (data.pageNumber > 1) {
            drawLetterhead(doc);
          }
        }
      });

      const finalY = (doc as any).lastAutoTable.finalY + 35;
      
      // Signature Section
      doc.setDrawColor(100, 100, 100);
      doc.setLineWidth(0.5);
      
      // Signature lines
      const signLineW = 70;
      doc.line(20, finalY, 20 + signLineW, finalY);
      doc.line(pageWidth - 20 - signLineW, finalY, pageWidth - 20, finalY);
      
      doc.setFontSize(8);
      doc.setTextColor(30, 41, 59);
      doc.setFont('helvetica', 'bold');
      doc.text('RESPONSÃVEL PELA ENTREGA', 20 + (signLineW/2), finalY + 5, { align: 'center' });
      doc.text('RESPONSÃVEL PELO SETOR (RECEBIMENTO)', pageWidth - 20 - (signLineW/2), finalY + 5, { align: 'center' });
      
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(7);
      const responsibleName = userProfile?.name || user?.displayName || 'ResponsÃ¡vel';
      doc.text(responsibleName, 20 + (signLineW/2), finalY + 10, { align: 'center' });
      doc.text(`Setor: ${data.sector}`, pageWidth - 20 - (signLineW/2), finalY + 10, { align: 'center' });

      // Disclaimer
      doc.setFontSize(8);
      doc.setTextColor(100, 100, 100);
      doc.text('Confirmo o recebimento dos materiais acima relacionados para uso exclusivo no setor designado.', pageWidth/2, finalY + 25, { align: 'center' });

      // Footer (Institutional Address from model)
      doc.setFontSize(7);
      doc.setTextColor(120, 113, 108);
      doc.setDrawColor(226, 232, 240);
      doc.line(14, pageHeight - 20, pageWidth - 14, pageHeight - 20);
      
      const footerLine1 = 'CEO - Centro de Especialidades OdontolÃ³gicas.';
      const footerLine2 = 'Fone: (88) 3614-3156 . Fax: (88) 3614-3245';
      doc.text(footerLine1, pageWidth / 2, pageHeight - 12, { align: 'center' });
      doc.text(footerLine2, pageWidth / 2, pageHeight - 8, { align: 'center' });

      const fileName = `RECIBO-${data.sector.toUpperCase().replace(/ /g, '-')}-${format(new Date(), 'ddMMyy-HHmm')}.pdf`;
      doc.save(fileName);
      showToast("Comprovante individual gerado com sucesso!", "success");
    } catch (error) {
      console.error("Receipt PDF Error:", error);
      showToast("Erro ao gerar PDF do comprovante", "error");
    }
  };


  const handleExportConsumptionPDF = () => {
    try {
      const doc = new jsPDF();
      const pageWidth = doc.internal.pageSize.width;
      
      // Minimalist Header (No heavy boxes)
      // Simulated Logo / Icon (Simple and clean)
      doc.setDrawColor(225, 29, 72); // rose-600 color for medical accent
      doc.setLineWidth(1.5);
      doc.line(14, 15, 24, 15); // Horizontal line of a plus
      doc.line(19, 10, 19, 20); // Vertical line of a plus
      
      doc.setFontSize(16);
      doc.setTextColor(28, 25, 23); // dark stone
      doc.setFont('helvetica', 'bold');
      doc.text('CEO', 28, 17);
      
      doc.setFontSize(8);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(120, 113, 108);
      doc.text('GESTÃƒO DE ALMOXARIFADO E FARMÃCIA', 28, 22);

      doc.setDrawColor(231, 229, 228); // light border
      doc.setLineWidth(0.5);
      doc.line(14, 28, pageWidth - 14, 28);
      
      // Title and Date
      doc.setFontSize(14);
      doc.setTextColor(28, 25, 23);
      doc.setFont('helvetica', 'bold');
      doc.text('RelatÃ³rio de Consumo por Setor', 14, 40);
      
      doc.setFontSize(9);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(120, 113, 108);
      doc.text(`PerÃ­odo: ${format(parseISO(reportRange.start), 'dd/MM/yyyy')} a ${format(parseISO(reportRange.end), 'dd/MM/yyyy')}`, 14, 46);
      doc.text(`Emitido em: ${format(new Date(), 'dd/MM/yyyy HH:mm')}`, 14, 51);
      
      // Summary Box (Minimalist) - Only for Admin
      if (isAdmin) {
        const totalValue = reportData.consumptionReport.reduce((sum, i) => sum + i.totalValue, 0);
        doc.setFillColor(250, 250, 249); // stone-50
        doc.roundedRect(pageWidth - 85, 35, 71, 18, 2, 2, 'F');
        doc.setFontSize(8);
        doc.setTextColor(120, 113, 108);
        doc.text('VALOR TOTAL CONSUMIDO', pageWidth - 80, 42);
        doc.setFontSize(11);
        doc.setTextColor(28, 25, 23);
        doc.setFont('helvetica', 'bold');
        doc.text(new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(totalValue), pageWidth - 80, 49);
      }

      // Table Data
      const tableData: any[] = [];
      reportData.consumptionBySector.forEach(sectorGroup => {
        // Sector Header
        const rowHeader: any[] = [
          { 
            content: sectorGroup.sector, 
            colSpan: isAdmin ? 4 : 3, 
            styles: { 
              fillColor: [250, 250, 249],
              textColor: [28, 25, 23], 
              fontStyle: 'bold',
              cellPadding: 4,
              fontSize: 10
            } 
          }
        ];

        if (isAdmin) {
          rowHeader.push({ 
            content: new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(sectorGroup.totalValue), 
            styles: { 
              fillColor: [250, 250, 249],
              halign: 'right', 
              fontStyle: 'bold' 
            } 
          });
        }
        
        tableData.push(rowHeader);
        
        // Items
        Object.values(sectorGroup.items).sort((a, b) => b.quantity - a.quantity).forEach(item => {
          const row: any[] = [
            { content: item.name, styles: { cellPadding: { left: 8 } } },
            item.category,
            { content: item.quantity.toString(), styles: { halign: 'center' } }
          ];

          if (isAdmin) {
            row.push({ content: new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(item.value), styles: { halign: 'right' } });
          }

          tableData.push(row);
        });
      });
      
      const headers = ['Item / Produto', 'Categoria', 'Qtd'];
      if (isAdmin) headers.push('Total (R$)');

      autoTable(doc, {
        startY: 60,
        head: [headers],
        body: tableData,
        theme: 'plain', 
        headStyles: { 
          textColor: [120, 113, 108], 
          fontSize: 8, 
          fontStyle: 'bold',
          halign: 'center',
          cellPadding: 4
        },
        styles: { 
          fontSize: 9, 
          cellPadding: 3,
          textColor: [68, 64, 60],
          lineWidth: 0 // Remove default borders
        },
        columnStyles: {
          0: { cellWidth: 'auto' },
          1: { cellWidth: 40 },
          2: { cellWidth: 20, halign: 'center' as any },
          3: { cellWidth: 35, halign: 'right' as any }
        },
        didParseCell: (data) => {
          if (data.section === 'body') {
            data.cell.styles.lineWidth = { bottom: 0.1 };
            data.cell.styles.lineColor = [231, 229, 228];
          }
          if (data.section === 'head') {
            data.cell.styles.lineWidth = { bottom: 0.5 };
            data.cell.styles.lineColor = [28, 25, 23];
          }
        },
        didDrawPage: (data) => {
          doc.setFontSize(7);
          doc.setTextColor(168, 162, 158);
          doc.text(`Documento emitido pelo Sistema de GestÃ£o Hospitalar - PÃ¡gina ${doc.getNumberOfPages()}`, 14, doc.internal.pageSize.height - 10);
        }
      });
      
      const fileName = `Relatorio_Consumo_CEO_${format(new Date(), 'dd-MM-yyyy')}.pdf`;
      doc.save(fileName);
      showToast("RelatÃ³rio profissional exportado!", "success");
    } catch (error) {
      console.error('Error exporting PDF:', error);
      showToast("Erro ao gerar PDF profissional.", "error");
    }
  };

  const reportData = useMemo(() => {
    const start = startOfDay(parseISO(reportRange.start));
    const end = endOfDay(parseISO(reportRange.end));
    const isAdmin = userProfile?.role === 'ADMIN' || 
                    user?.email === 'gerlianemagalhaes79@gmail.com' || 
                    userProfile?.sector === 'Almoxarifado';
    const effectiveSectorFilter = isAdmin ? reportSectorFilter : (selectedSector || 'none');

    const filteredTrans = transactions.filter(t => {
      if (t.deletedAt) return false;
      const d = new Date(t.date);
      const inRange = d >= start && d <= end;
    const matchesSector = effectiveSectorFilter === 'all' || 
                          t.sector === effectiveSectorFilter || 
                          (effectiveSectorFilter === 'FarmÃ¡cia' && t.sector === 'FarmÃ¡cia (Consumo Interno)');
      return inRange && matchesSector;
    });

    const regularEntriesTrans = filteredTrans.filter(t => t.type === 'entry' && !t.isReturn);
    const returnTrans = filteredTrans.filter(t => t.type === 'entry' && t.isReturn === true);
    const exitTrans = filteredTrans.filter(t => t.type === 'exit');

    const entries = regularEntriesTrans.reduce((sum, t) => sum + t.quantity, 0);
    const exits = exitTrans.reduce((sum, t) => sum + t.quantity, 0) - returnTrans.reduce((sum, t) => sum + t.quantity, 0);
    
    const entriesValue = regularEntriesTrans.reduce((sum, t) => {
      const item = items.find(i => i.id === t.item_id);
      return sum + (t.quantity * (Number(item?.unit_price) || 0));
    }, 0);
    
    const exitsValue = exitTrans.reduce((sum, t) => {
      const item = items.find(i => i.id === t.item_id);
      return sum + (t.quantity * (Number(item?.unit_price) || 0));
    }, 0) - returnTrans.reduce((sum, t) => {
      const item = items.find(i => i.id === t.item_id);
      return sum + (t.quantity * (Number(item?.unit_price) || 0));
    }, 0);

    // Extra vs Contract stats
    const originStats = {
      extra: { entries: 0, exits: 0, current: 0 },
      contract: { entries: 0, exits: 0, current: 0 },
      donation: { entries: 0, exits: 0, current: 0 }
    };

    filteredTrans.forEach(t => {
      const origin = t.origin || 'contract';
      if (t.type === 'entry') {
        if (t.isReturn) {
          originStats[origin].exits -= t.quantity;
        } else {
          originStats[origin].entries += t.quantity;
        }
      } else {
        originStats[origin].exits += t.quantity;
      }
    });

    const filteredItems = items.filter(item => {
      if (item.deletedAt) return false;
      
      // If not admin, only see items from their own location
      if (!isAdmin) {
        const userLocation = userProfile?.sector === 'FarmÃ¡cia' ? 'FarmÃ¡cia' : 'Almoxarifado';
        return (item.location || 'Almoxarifado') === userLocation;
      }
      
      // If admin, respect the sector filter if it maps to a location
      if (reportSectorFilter === 'FarmÃ¡cia') {
        return item.location === 'FarmÃ¡cia';
      } else if (reportSectorFilter === 'Almoxarifado') {
        return (item.location || 'Almoxarifado') === 'Almoxarifado';
      }
      
      // If 'all' or other sector, show everything for admin
      return true;
    });

    filteredItems.forEach(item => {
      const origin = item.origin || 'contract';
      originStats[origin].current += (Number(item.quantity) || 0);
    });

    // Group by date for line chart
    const dailyData: Record<string, { date: string, entries: number, exits: number, sortKey: string }> = {};
    filteredTrans.forEach(t => {
      const dateObj = new Date(t.date);
      const dateKey = format(dateObj, 'dd/MM');
      const sortKey = format(dateObj, 'yyyy-MM-dd');
      if (!dailyData[sortKey]) dailyData[sortKey] = { date: dateKey, entries: 0, exits: 0, sortKey };
      if (t.type === 'entry') {
        if (t.isReturn) {
          dailyData[sortKey].exits -= t.quantity;
        } else {
          dailyData[sortKey].entries += t.quantity;
        }
      } else {
        dailyData[sortKey].exits += t.quantity;
      }
    });

    // Group by category for pie chart (quantity)
    const categoryData: Record<string, number> = {};
    // Group by category for value chart
    const categoryValueData: Record<string, number> = {};
    
    const filteredItemsForValue = items.filter(item => {
      if (item.deletedAt) return false;
      
      // If not admin, only see items from their own location
      if (!isAdmin) {
        const userLocation = userProfile?.sector === 'FarmÃ¡cia' ? 'FarmÃ¡cia' : 'Almoxarifado';
        return (item.location || 'Almoxarifado') === userLocation;
      }
      
      // If admin, respect the sector filter if it maps to a location
      if (reportSectorFilter === 'FarmÃ¡cia') {
        return item.location === 'FarmÃ¡cia';
      } else if (reportSectorFilter === 'Almoxarifado') {
        return (item.location || 'Almoxarifado') === 'Almoxarifado';
      }
      
      return true;
    });

    filteredItemsForValue.forEach(item => {
      const cat = item.category || 'Outros';
      const qty = Number(item.quantity) || 0;
      const price = Number(item.unit_price) || 0;
      categoryData[cat] = (categoryData[cat] || 0) + qty;
      categoryValueData[cat] = (categoryValueData[cat] || 0) + (qty * price);
    });

    // Group by sector for bar chart (stacked by category)
    const sectorData: Record<string, any> = {};
    const categoriesInSector: Set<string> = new Set();

    filteredTrans.filter(t => t.type === 'exit' && t.sector).forEach(t => {
      const item = items.find(i => i.id === t.item_id);
      const category = item?.category || 'Outros';
      categoriesInSector.add(category);
      
      const sectorKey = (t.sector === 'FarmÃ¡cia (Consumo Interno)') ? 'FarmÃ¡cia' : t.sector!;
      
      if (!sectorData[sectorKey]) {
        sectorData[sectorKey] = { name: sectorKey };
      }
      sectorData[sectorKey][category] = (sectorData[sectorKey][category] || 0) + t.quantity;
    });

    filteredTrans.filter(t => t.type === 'entry' && t.isReturn && t.sector).forEach(t => {
      const item = items.find(i => i.id === t.item_id);
      const category = item?.category || 'Outros';
      categoriesInSector.add(category);
      
      const sectorKey = (t.sector === 'FarmÃ¡cia (Consumo Interno)') ? 'FarmÃ¡cia' : t.sector!;
      
      if (!sectorData[sectorKey]) {
        sectorData[sectorKey] = { name: sectorKey };
      }
      sectorData[sectorKey][category] = (sectorData[sectorKey][category] || 0) - t.quantity;
    });

    // Consumption report with sector breakdown
    const consumptionReport: Record<string, { 
      name: string, 
      totalQuantity: number, 
      totalValue: number, 
      category: string, 
      supplier: string,
      sectors: Record<string, number>
    }> = {};

    // Consumption report grouped by sector
    const consumptionBySector: Record<string, {
      sector: string,
      totalValue: number,
      items: Record<string, {
        name: string,
        quantity: number,
        value: number,
        category: string
      }>
    }> = {};

    // Process exits
    filteredTrans.filter(t => t.type === 'exit').forEach(t => {
      const item = items.find(i => i.id === t.item_id);
      const price = Number(item?.unit_price) || 0;
      const value = t.quantity * price;
      let sector = t.sector || 'NÃ£o Informado';
      if (sector === 'FarmÃ¡cia (Consumo Interno)') sector = 'FarmÃ¡cia';
      
      if (!consumptionReport[t.item_name]) {
        consumptionReport[t.item_name] = { 
          name: t.item_name, 
          totalQuantity: 0, 
          totalValue: 0, 
          category: item?.category || 'Outros',
          supplier: item?.supplier || 'N/A',
          sectors: {}
        };
      }
      consumptionReport[t.item_name].totalQuantity += t.quantity;
      consumptionReport[t.item_name].totalValue += value;
      consumptionReport[t.item_name].sectors[sector] = (consumptionReport[t.item_name].sectors[sector] || 0) + t.quantity;

      // Group by Sector
      if (!consumptionBySector[sector]) {
        consumptionBySector[sector] = {
          sector,
          totalValue: 0,
          items: {}
        };
      }
      
      if (!consumptionBySector[sector].items[t.item_name]) {
        consumptionBySector[sector].items[t.item_name] = {
          name: t.item_name,
          quantity: 0,
          value: 0,
          category: item?.category || 'Outros'
        };
      }
      
      consumptionBySector[sector].totalValue += value;
      consumptionBySector[sector].items[t.item_name].quantity += t.quantity;
      consumptionBySector[sector].items[t.item_name].value += value;
    });

    // Subtract returns
    filteredTrans.filter(t => t.type === 'entry' && t.isReturn).forEach(t => {
      const item = items.find(i => i.id === t.item_id);
      const price = Number(item?.unit_price) || 0;
      const value = t.quantity * price;
      let sector = t.sector || 'NÃ£o Informado';
      if (sector === 'FarmÃ¡cia (Consumo Interno)') sector = 'FarmÃ¡cia';
      
      if (!consumptionReport[t.item_name]) {
        consumptionReport[t.item_name] = { 
          name: t.item_name, 
          totalQuantity: 0, 
          totalValue: 0, 
          category: item?.category || 'Outros',
          supplier: item?.supplier || 'N/A',
          sectors: {}
        };
      }
      consumptionReport[t.item_name].totalQuantity -= t.quantity;
      consumptionReport[t.item_name].totalValue -= value;
      consumptionReport[t.item_name].sectors[sector] = (consumptionReport[t.item_name].sectors[sector] || 0) - t.quantity;

      // Group by Sector
      if (!consumptionBySector[sector]) {
        consumptionBySector[sector] = {
          sector,
          totalValue: 0,
          items: {}
        };
      }
      
      if (!consumptionBySector[sector].items[t.item_name]) {
        consumptionBySector[sector].items[t.item_name] = {
          name: t.item_name,
          quantity: 0,
          value: 0,
          category: item?.category || 'Outros'
        };
      }
      
      consumptionBySector[sector].totalValue -= value;
      consumptionBySector[sector].items[t.item_name].quantity -= t.quantity;
      consumptionBySector[sector].items[t.item_name].value -= value;
    });

    // Group by supplier for value chart
    const supplierData: Record<string, number> = {};
    filteredItemsForValue.forEach(item => {
      const sup = item.supplier || 'Sem Fornecedor';
      const qty = Number(item.quantity) || 0;
      const price = Number(item.unit_price) || 0;
      supplierData[sup] = (supplierData[sup] || 0) + (qty * price);
    });

    const totalValue = filteredItemsForValue.reduce((sum, item) => sum + ((Number(item.quantity) || 0) * (Number(item.unit_price) || 0)), 0);

    // Most requested items
    const mostRequested: Record<string, number> = {};
    allRequestItems.forEach(ri => {
      const request = requests.find(r => r.id === ri.request_id);
      if (!request) return;
      
      // If not admin, only count items from their own sector
      if (!isAdmin && request.sector !== selectedSector) return;
      
      // If admin and sector filter is active, filter by that sector
      if (isAdmin && reportSectorFilter !== 'all' && request.sector !== reportSectorFilter) return;

      mostRequested[ri.product_name] = (mostRequested[ri.product_name] || 0) + ri.quantity_requested;
    });
    const topRequested = Object.entries(mostRequested)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 10);

    // Exits by reason
    const exitsByReason: Record<string, number> = {
      'consumo': 0,
      'doacao': 0,
      'vencido': 0,
      'perda': 0
    };
    filteredTrans.filter(t => t.type === 'exit').forEach(t => {
      const reason = t.exitReason || 'consumo';
      if (exitsByReason[reason] !== undefined) {
        exitsByReason[reason] += t.quantity;
      }
    });

    // Returns by Sector calculation
    const returnsBySectorMap: Record<string, { name: string; quantity: number; value: number }> = {};
    const returnsByReasonMap: Record<string, number> = {};

    filteredTrans.filter(t => t.type === 'entry' && t.isReturn).forEach(t => {
      const item = items.find(i => i.id === t.item_id);
      const price = Number(item?.unit_price) || 0;
      const val = t.quantity * price;
      let sec = t.sector || 'NÃ£o Informado';
      if (sec === 'FarmÃ¡cia (Consumo Interno)') sec = 'FarmÃ¡cia';

      if (!returnsBySectorMap[sec]) {
        returnsBySectorMap[sec] = { name: sec, quantity: 0, value: 0 };
      }
      returnsBySectorMap[sec].quantity += t.quantity;
      returnsBySectorMap[sec].value += val;

      const reason = t.returnReason || 'NÃ£o especificado';
      returnsByReasonMap[reason] = (returnsByReasonMap[reason] || 0) + t.quantity;
    });

    const returnsBySector = Object.values(returnsBySectorMap).sort((a, b) => b.quantity - a.quantity);
    const returnsByReason = Object.entries(returnsByReasonMap)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);
    const totalReturnsCount = filteredTrans.filter(t => t.type === 'entry' && t.isReturn).reduce((sum, t) => sum + t.quantity, 0);
    const totalReturnsValue = filteredTrans.filter(t => t.type === 'entry' && t.isReturn).reduce((sum, t) => {
      const item = items.find(i => i.id === t.item_id);
      return sum + (t.quantity * (Number(item?.unit_price) || 0));
    }, 0);

    return {
      entries,
      exits,
      entriesValue,
      exitsValue,
      daily: Object.values(dailyData).sort((a, b) => a.sortKey.localeCompare(b.sortKey)),
      categories: Object.entries(categoryData)
        .map(([name, value]) => ({ name, value }))
        .filter(c => c.value > 0)
        .sort((a, b) => b.value - a.value),
      consumptionCategories: Object.entries(
        (() => {
          const acc: Record<string, number> = {};
          filteredTrans.forEach(t => {
            const item = items.find(i => i.id === t.item_id);
            const cat = item?.category || 'Outros';
            if (t.type === 'exit') {
              acc[cat] = (acc[cat] || 0) + t.quantity;
            } else if (t.type === 'entry' && t.isReturn) {
              acc[cat] = (acc[cat] || 0) - t.quantity;
            }
          });
          return acc;
        })()
      ).map(([name, value]) => ({ name, value })),
      categoryValues: Object.entries(categoryValueData)
        .map(([name, value]) => ({ name, value }))
        .filter(c => c.value > 0)
        .sort((a, b) => b.value - a.value),
      sectors: Object.values(sectorData),
      categoriesInSector: Array.from(categoriesInSector),
      suppliers: Object.entries(supplierData)
        .map(([name, value]) => ({ name, value }))
        .filter(s => s.value > 0)
        .sort((a, b) => b.value - a.value),
      consumptionReport: Object.values(consumptionReport).sort((a, b) => b.totalValue - a.totalValue),
      consumptionBySector: Object.values(consumptionBySector).sort((a, b) => b.totalValue - a.totalValue),
      totalValue,
      originStats,
      topRequested,
      topConsumed: Object.values(consumptionReport)
        .map(i => ({ name: i.name, value: i.totalQuantity }))
        .sort((a, b) => b.value - a.value)
        .slice(0, 10),
      exitsByReason,
      returnsBySector,
      returnsByReason,
      totalReturnsCount,
      totalReturnsValue
    };
  }, [transactions, items, reportRange, reportSectorFilter, allRequestItems, requests, userProfile, isAdmin, selectedSector]);

  const categoryDistribution = useMemo(() => {
    const map: Record<string, { productNames: Set<string>; totalQty: number; value: number }> = {};
    items
      .filter(i => !i.deletedAt && i.quantity > 0 && (i.location || 'Almoxarifado') === inventoryLocation)
      .forEach(i => {
        const cat = i.category || 'Geral';
        if (!map[cat]) map[cat] = { productNames: new Set<string>(), totalQty: 0, value: 0 };
        map[cat].productNames.add(i.name.trim());
        map[cat].totalQty += (Number(i.quantity) || 0);
        map[cat].value += ((Number(i.quantity) || 0) * (Number(i.unit_price) || 0));
      });

    const entries = Object.entries(map).map(([category, data]) => ({
      category,
      count: data.productNames.size,
      totalQty: data.totalQty,
      value: data.value,
    }));

    const maxCount = Math.max(...entries.map(m => m.count), 1);
    const maxQty = Math.max(...entries.map(m => m.totalQty), 1);

    return entries
      .map(data => ({
        ...data,
        typePercentage: Math.min(100, Math.round((data.count / maxCount) * 100)),
        unitPercentage: Math.min(100, Math.round((data.totalQty / maxQty) * 100))
      }))
      .sort((a, b) => distribViewMode === 'types' ? b.count - a.count : b.totalQty - a.totalQty)
      .slice(0, 12);
  }, [items, inventoryLocation, distribViewMode]);

  if (loading) {
    return (
      <div className="min-h-screen bg-[#F5F5F4] flex items-center justify-center">
        <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-[#1C1917]"></div>
      </div>
    );
  }

  if (!user) {
    const loginLogo = appRectangularLogo || appLogo;
    return (
      <div className="min-h-screen bg-slate-100 flex items-center justify-center p-6">
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-white p-10 rounded-[40px] shadow-2xl max-w-md w-full border border-slate-200"
        >
          <div className="text-center mb-8">
            {loginLogo ? (
              <div className="w-full max-w-[260px] h-24 rounded-2xl overflow-hidden bg-white border border-blue-200/80 p-2.5 shadow-md mx-auto mb-6 flex items-center justify-center ring-4 ring-blue-500/10">
                <img src={loginLogo} alt="Logo CEO" className="max-w-full max-h-full object-contain" />
              </div>
            ) : (
              <div className="bg-gradient-to-br from-blue-700 via-blue-800 to-indigo-900 w-24 h-24 rounded-3xl flex items-center justify-center mx-auto mb-6 shadow-xl overflow-hidden border-4 border-white ring-4 ring-blue-500/10 text-white">
                <Package className="w-12 h-12" />
              </div>
            )}
            <div className="mb-4 text-center">
              <h1 className="text-3xl font-black tracking-tight text-slate-900 uppercase">
                CEO
              </h1>
            </div>
            <div className="h-0.5 w-12 bg-blue-100 mx-auto mb-4 rounded-full" />
            <p className="text-slate-500 text-[10px] font-black uppercase tracking-widest bg-blue-50 border border-blue-100/80 px-3 py-1 rounded-full w-fit mx-auto">
              Almoxarifado Inteligente
            </p>
          </div>

          <div className="space-y-6">
            <button 
              onClick={handleGoogleLogin}
              disabled={loginLoading}
              className="w-full bg-white border border-slate-200 text-slate-800 py-4 rounded-2xl font-extrabold flex items-center justify-center gap-3 hover:bg-slate-50 hover:border-blue-300 transition-all shadow-sm active:scale-[0.98] disabled:opacity-50 group"
            >
              {loginLoading ? (
                <div className="animate-spin rounded-full h-5 w-5 border-b-2 border-blue-700"></div>
              ) : (
                <>
                  <img src="https://www.gstatic.com/firebasejs/ui/2.0.0/images/auth/google.svg" className="w-5 h-5 group-hover:scale-110 transition-transform" alt="Google" />
                  <span>Entrar com Google</span>
                </>
              )}
            </button>
            <p className="text-[10px] text-slate-400 text-center font-extrabold uppercase tracking-widest mt-4">
              Apenas e-mails autorizados pelo administrador
            </p>
          </div>

          <div className="mt-8 text-center pt-4 border-t border-slate-100">
            <p className="text-[10px] text-slate-400 uppercase tracking-widest font-extrabold">Acesso restrito a funcionÃ¡rios autorizados</p>
          </div>
        </motion.div>
      </div>
    );
  }

  const isExpired = (item: Item) => {
    if (item.quantity <= 0) return false;
    const dateStr = item.expiry_date;
    if (!dateStr || dateStr === 'Indeterminada') return false;
    const expiry = new Date(dateStr);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return expiry < today;
  };

  const isNearExpiry = (item: Item) => {
    if (item.quantity <= 0) return false;
    const dateStr = item.expiry_date;
    if (!dateStr || dateStr === 'Indeterminada') return false;
    const expiry = new Date(dateStr);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const oneMonthFromNow = new Date();
    oneMonthFromNow.setMonth(today.getMonth() + 1);
    return expiry >= today && expiry <= oneMonthFromNow;
  };

  const filteredItems = items.filter(i => {
    const normalizedSearch = normalizeString(searchTerm);
    const itemLocation = i.location || 'Almoxarifado';
    return !i.deletedAt && 
    i.quantity > 0 && 
    itemLocation === inventoryLocation &&
    ((normalizeString(i.name).includes(normalizedSearch) || 
    normalizeString(i.supplier).includes(normalizedSearch) ||
    normalizeString(i.category).includes(normalizedSearch) ||
    normalizeString(i.batch_number).includes(normalizedSearch)) &&
    (originFilter === 'all' || i.origin === originFilter) &&
    (categoryFilter === 'all' || i.category === categoryFilter));
  });

  const groupedItems = items.filter(i => !i.deletedAt && i.quantity > 0 && (i.location || 'Almoxarifado') === inventoryLocation).reduce((acc, item) => {
    if (!acc[item.name]) {
      const weeklyExitRate = weeklyExitRates[item.name] || 0;
      
      acc[item.name] = {
        name: item.name,
        total_quantity: 0,
        min_quantity: weeklyExitRate > 0 ? Math.ceil(weeklyExitRate * 5) : item.min_quantity,
        category: item.category,
        supplier: item.supplier,
        unit_measure: item.unit_measure || null,
        batches: [],
        weeklyExitRate: weeklyExitRate,
        durationWeeks: 0
      };
    }
    acc[item.name].total_quantity += item.quantity;
    if (!acc[item.name].unit_measure && item.unit_measure) {
      acc[item.name].unit_measure = item.unit_measure;
    }
    acc[item.name].batches.push(item);
    
    // Update duration
    if (acc[item.name].weeklyExitRate > 0) {
      acc[item.name].durationWeeks = acc[item.name].total_quantity / acc[item.name].weeklyExitRate;
    } else {
      acc[item.name].durationWeeks = 'infinite';
    }
    
    return acc;
  }, {} as Record<string, ItemGroup>);

  const lowStockItems = Object.values(groupedItems).filter(group => 
    group.total_quantity <= group.min_quantity
  );

  const expiredItems = items.filter(i => !i.deletedAt && (i.location || 'Almoxarifado') === inventoryLocation && isExpired(i));
  const nearExpiryItems = items.filter(i => !i.deletedAt && (i.location || 'Almoxarifado') === inventoryLocation && isNearExpiry(i));
  const totalAlertsCount = lowStockItems.length + expiredItems.length + nearExpiryItems.length;
  const totalVolume = items
    .filter(i => !i.deletedAt && (i.location || 'Almoxarifado') === inventoryLocation)
    .reduce((sum, item) => sum + (Number(item.quantity) || 0), 0);
  const totalInventoryValue = items
    .filter(i => !i.deletedAt && (i.location || 'Almoxarifado') === inventoryLocation)
    .reduce((sum, item) => sum + ((Number(item.quantity) || 0) * (Number(item.unit_price) || 0)), 0);

  const recentTransactions = transactions
    .filter(t => (t.location || 'Almoxarifado') === inventoryLocation)
    .slice(0, 5);

  const pendingRequestsCount = requests.filter(r => 
    !r.deletedAt && 
    (r.status === 'PENDENTE' || r.status === 'EM_SEPARACAO' || r.status === 'DEVOLUCAO_PENDENTE') &&
    (isAdmin ? true : r.sector === selectedSector)
  ).length;

  const groupedArray: ItemGroup[] = (Object.values(groupedItems) as ItemGroup[])
    .filter(group => {
      // Apply search and filters to the grouped items for the inventory list
      const normalizedSearch = normalizeString(searchTerm);
      const matchesSearch = normalizeString(group.name).includes(normalizedSearch) ||
                           normalizeString(group.supplier).includes(normalizedSearch) ||
                           normalizeString(group.category).includes(normalizedSearch);
      
      const matchesOrigin = originFilter === 'all' || group.batches.some(b => b.origin === originFilter);
      const matchesCategory = categoryFilter === 'all' || group.category === categoryFilter;
      
      return matchesSearch && matchesOrigin && matchesCategory;
    })
    .sort((a, b) => {
      if (inventorySort === 'name_asc') {
        return a.name.localeCompare(b.name);
      } else if (inventorySort === 'name_desc') {
        return b.name.localeCompare(a.name);
      } else if (inventorySort === 'duration_asc') {
        const durA = a.durationWeeks === 'infinite' ? Number.MAX_SAFE_INTEGER : a.durationWeeks;
        const durB = b.durationWeeks === 'infinite' ? Number.MAX_SAFE_INTEGER : b.durationWeeks;
        return durA - durB;
      } else {
        const durA = a.durationWeeks === 'infinite' ? Number.MAX_SAFE_INTEGER : a.durationWeeks;
        const durB = b.durationWeeks === 'infinite' ? Number.MAX_SAFE_INTEGER : b.durationWeeks;
        return durB - durA;
      }
    });

  return (
    <div className="min-h-screen bg-[#F5F5F4] text-[#1C1917] font-sans">
      {/* Mobile Header */}
      <header className="lg:hidden fixed top-0 left-0 right-0 h-16 bg-slate-900 border-b border-slate-800 flex items-center justify-between px-4 z-20 text-white shadow-md">
        <div className="flex items-center gap-2.5">
          {appLogo ? (
            <div className="w-9 h-9 rounded-xl overflow-hidden bg-white p-0.5 border border-slate-700 shadow-sm flex items-center justify-center shrink-0">
              <img src={appLogo} alt="Logo CEO" className="w-full h-full object-contain" />
            </div>
          ) : (
            <div className="bg-gradient-to-br from-blue-600 to-indigo-800 p-2 rounded-xl text-white shadow-sm shrink-0">
              <Package className="w-5 h-5" />
            </div>
          )}
          <div className="flex flex-col min-w-0">
            <h1 className="font-black text-xl tracking-tight text-white leading-none">CEO</h1>
            <span className="text-[10px] font-black text-blue-300 tracking-wider uppercase mt-0.5">Almoxarifado</span>
          </div>
        </div>
        <button 
          onClick={() => setIsMobileMenuOpen(!isMobileMenuOpen)}
          className="p-2 hover:bg-slate-800 rounded-xl transition-colors text-slate-300"
        >
          {isMobileMenuOpen ? <X size={22} /> : <Menu size={22} />}
        </button>
      </header>

      {/* Sidebar Overlay */}
      <AnimatePresence>
        {isMobileMenuOpen && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            onClick={() => setIsMobileMenuOpen(false)}
            className="fixed inset-0 bg-slate-950/40 backdrop-blur-sm z-30 lg:hidden"
          />
        )}
      </AnimatePresence>

      {/* Sidebar */}
      <aside className={`fixed lg:left-0 top-0 h-full w-64 bg-white border-r border-blue-100/80 p-5 flex flex-col gap-6 z-40 shadow-sm transition-transform duration-300 transform ${isMobileMenuOpen ? 'translate-x-0' : '-translate-x-full lg:translate-x-0'}`}>
        <div className="flex items-center justify-between lg:justify-start gap-3 px-1 pt-1">
          <div className="flex items-center gap-3">
            {appLogo ? (
              <div className="w-11 h-11 rounded-2xl overflow-hidden bg-white border border-blue-200/80 p-1 shadow-md shadow-blue-500/10 flex items-center justify-center shrink-0 ring-2 ring-blue-500/10">
                <img src={appLogo} alt="Logo CEO" className="w-full h-full object-contain" />
              </div>
            ) : (
              <div className="bg-gradient-to-br from-blue-700 via-blue-800 to-indigo-900 p-2.5 rounded-2xl shadow-md shadow-blue-500/20 text-white ring-2 ring-blue-500/20 shrink-0">
                <Package className="w-6 h-6" />
              </div>
            )}
            <div className="flex flex-col min-w-0">
              <h1 className="font-black text-2xl tracking-tight text-slate-900 leading-none">CEO</h1>
              <span className="text-[10px] font-black text-blue-700 bg-blue-50 border border-blue-100 px-2 py-0.5 rounded-md uppercase tracking-widest leading-none block w-fit mt-1">
                Almoxarifado
              </span>
            </div>
          </div>
          <button 
            onClick={() => setIsMobileMenuOpen(false)}
            className="lg:hidden p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded-xl"
          >
            <X size={20} />
          </button>
        </div>

        <nav className="flex flex-col gap-1.5 overflow-y-auto pr-0.5">
          {userProfile && (
            <>
              {(isAdmin || userProfile.role === 'ADMIN') ? (
                <>
                  <button 
                    onClick={() => { setActiveTab('dashboard'); setIsMobileMenuOpen(false); }}
                    className={`group flex items-center justify-between px-3.5 py-2.5 rounded-2xl transition-all duration-200 text-xs ${
                      activeTab === 'dashboard' 
                        ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white font-extrabold shadow-md shadow-blue-600/20' 
                        : 'text-slate-600 hover:text-blue-700 hover:bg-blue-50/80 font-bold'
                    }`}
                  >
                    <div className="flex items-center gap-3">
                      <LayoutDashboard size={18} className={activeTab === 'dashboard' ? 'text-white' : 'text-slate-400 group-hover:text-blue-600 transition-colors'} />
                      <span>Dashboard</span>
                    </div>
                  </button>

                  <button 
                    onClick={() => { setActiveTab('inventory'); setIsMobileMenuOpen(false); }}
                    className={`group flex items-center justify-between px-3.5 py-2.5 rounded-2xl transition-all duration-200 text-xs ${
                      activeTab === 'inventory' 
                        ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white font-extrabold shadow-md shadow-blue-600/20' 
                        : 'text-slate-600 hover:text-blue-700 hover:bg-blue-50/80 font-bold'
                    }`}
                  >
                    <div className="flex items-center gap-3">
                      <Package size={18} className={activeTab === 'inventory' ? 'text-white' : 'text-slate-400 group-hover:text-blue-600 transition-colors'} />
                      <span>Estoque</span>
                    </div>
                  </button>

                  <button 
                    onClick={() => { setActiveTab('history'); setIsMobileMenuOpen(false); }}
                    className={`group flex items-center justify-between px-3.5 py-2.5 rounded-2xl transition-all duration-200 text-xs ${
                      activeTab === 'history' 
                        ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white font-extrabold shadow-md shadow-blue-600/20' 
                        : 'text-slate-600 hover:text-blue-700 hover:bg-blue-50/80 font-bold'
                    }`}
                  >
                    <div className="flex items-center gap-3">
                      <History size={18} className={activeTab === 'history' ? 'text-white' : 'text-slate-400 group-hover:text-blue-600 transition-colors'} />
                      <span>HistÃ³rico</span>
                    </div>
                  </button>

                  <button 
                    onClick={() => { setActiveTab('requests'); setIsMobileMenuOpen(false); }}
                    className={`group flex items-center justify-between px-3.5 py-2.5 rounded-2xl transition-all duration-200 text-xs ${
                      activeTab === 'requests' 
                        ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white font-extrabold shadow-md shadow-blue-600/20' 
                        : 'text-slate-600 hover:text-blue-700 hover:bg-blue-50/80 font-bold'
                    }`}
                  >
                    <div className="flex items-center gap-3">
                      <FileText size={18} className={activeTab === 'requests' ? 'text-white' : 'text-slate-400 group-hover:text-blue-600 transition-colors'} />
                      <span>SolicitaÃ§Ãµes</span>
                    </div>
                    {pendingRequestsCount > 0 && (
                      <span className={`text-[10px] font-black px-2 py-0.5 rounded-full ${
                        activeTab === 'requests' ? 'bg-amber-400 text-slate-950' : 'bg-sky-100 text-sky-800'
                      }`}>
                        {pendingRequestsCount}
                      </span>
                    )}
                  </button>

                  <button 
                    onClick={() => { setActiveTab('admin-devolutions'); setIsMobileMenuOpen(false); }}
                    className={`group flex items-center justify-between px-3.5 py-2.5 rounded-2xl transition-all duration-200 text-xs ${
                      activeTab === 'admin-devolutions' 
                        ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white font-extrabold shadow-md shadow-blue-600/20' 
                        : 'text-slate-600 hover:text-blue-700 hover:bg-blue-50/80 font-bold'
                    }`}
                  >
                    <div className="flex items-center gap-3">
                      <RotateCcw size={18} className={activeTab === 'admin-devolutions' ? 'text-white' : 'text-slate-400 group-hover:text-blue-600 transition-colors'} />
                      <span>DevoluÃ§Ãµes</span>
                    </div>
                  </button>

                  <button 
                    onClick={() => { setActiveTab('trash'); setIsMobileMenuOpen(false); }}
                    className={`group flex items-center justify-between px-3.5 py-2.5 rounded-2xl transition-all duration-200 text-xs ${
                      activeTab === 'trash' 
                        ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white font-extrabold shadow-md shadow-blue-600/20' 
                        : 'text-slate-600 hover:text-blue-700 hover:bg-blue-50/80 font-bold'
                    }`}
                  >
                    <div className="flex items-center gap-3">
                      <Trash2 size={18} className={activeTab === 'trash' ? 'text-white' : 'text-slate-400 group-hover:text-blue-600 transition-colors'} />
                      <span>Lixeira</span>
                    </div>
                  </button>

                  <button 
                    onClick={() => { setActiveTab('reports'); setIsMobileMenuOpen(false); }}
                    className={`group flex items-center justify-between px-3.5 py-2.5 rounded-2xl transition-all duration-200 text-xs ${
                      activeTab === 'reports' 
                        ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white font-extrabold shadow-md shadow-blue-600/20' 
                        : 'text-slate-600 hover:text-blue-700 hover:bg-blue-50/80 font-bold'
                    }`}
                  >
                    <div className="flex items-center gap-3">
                      <BarChart3 size={18} className={activeTab === 'reports' ? 'text-white' : 'text-slate-400 group-hover:text-blue-600 transition-colors'} />
                      <span>RelatÃ³rios</span>
                    </div>
                  </button>

                  <button 
                    onClick={() => { setActiveTab('users'); setIsMobileMenuOpen(false); }}
                    className={`group flex items-center justify-between px-3.5 py-2.5 rounded-2xl transition-all duration-200 text-xs ${
                      activeTab === 'users' 
                        ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white font-extrabold shadow-md shadow-blue-600/20' 
                        : 'text-slate-600 hover:text-blue-700 hover:bg-blue-50/80 font-bold'
                    }`}
                  >
                    <div className="flex items-center gap-3">
                      <Users size={18} className={activeTab === 'users' ? 'text-white' : 'text-slate-400 group-hover:text-blue-600 transition-colors'} />
                      <span>UsuÃ¡rios</span>
                    </div>
                  </button>
                </>
              ) : (
                <>
                  <button 
                    onClick={() => { setActiveTab('inventory'); setIsMobileMenuOpen(false); }}
                    className={`group flex items-center justify-between px-3.5 py-2.5 rounded-2xl transition-all duration-200 text-xs ${
                      activeTab === 'inventory' 
                        ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white font-extrabold shadow-md shadow-blue-600/20' 
                        : 'text-slate-600 hover:text-blue-700 hover:bg-blue-50/80 font-bold'
                    }`}
                  >
                    <div className="flex items-center gap-3">
                      <Package size={18} className={activeTab === 'inventory' ? 'text-white' : 'text-slate-400 group-hover:text-blue-600 transition-colors'} />
                      <span>{selectedSector === 'FarmÃ¡cia' || userProfile?.sector === 'FarmÃ¡cia' ? 'Estoque da FarmÃ¡cia' : 'Estoque'}</span>
                    </div>
                  </button>

                  <button 
                    onClick={() => { setActiveTab('new-request'); setIsMobileMenuOpen(false); }}
                    className={`group flex items-center justify-between px-3.5 py-2.5 rounded-2xl transition-all duration-200 text-xs ${
                      activeTab === 'new-request' 
                        ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white font-extrabold shadow-md shadow-blue-600/20' 
                        : 'text-slate-600 hover:text-blue-700 hover:bg-blue-50/80 font-bold'
                    }`}
                  >
                    <div className="flex items-center gap-3">
                      <Plus size={18} className={activeTab === 'new-request' ? 'text-white' : 'text-slate-400 group-hover:text-blue-600 transition-colors'} />
                      <span>Nova SolicitaÃ§Ã£o</span>
                    </div>
                  </button>

                  <button 
                    onClick={() => { setActiveTab('devolution'); setIsMobileMenuOpen(false); }}
                    className={`group flex items-center justify-between px-3.5 py-2.5 rounded-2xl transition-all duration-200 text-xs ${
                      activeTab === 'devolution' 
                        ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white font-extrabold shadow-md shadow-blue-600/20' 
                        : 'text-slate-600 hover:text-blue-700 hover:bg-blue-50/80 font-bold'
                    }`}
                  >
                    <div className="flex items-center gap-3">
                      <RotateCcw size={18} className={activeTab === 'devolution' ? 'text-white' : 'text-slate-400 group-hover:text-blue-600 transition-colors'} />
                      <span>DevoluÃ§Ã£o de Materiais</span>
                    </div>
                  </button>

                  <button 
                    onClick={() => { setActiveTab('my-requests'); setIsMobileMenuOpen(false); }}
                    className={`group flex items-center justify-between px-3.5 py-2.5 rounded-2xl transition-all duration-200 text-xs ${
                      activeTab === 'my-requests' 
                        ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white font-extrabold shadow-md shadow-blue-600/20' 
                        : 'text-slate-600 hover:text-blue-700 hover:bg-blue-50/80 font-bold'
                    }`}
                  >
                    <div className="flex items-center gap-3">
                      <FileText size={18} className={activeTab === 'my-requests' ? 'text-white' : 'text-slate-400 group-hover:text-blue-600 transition-colors'} />
                      <span>Minhas SolicitaÃ§Ãµes</span>
                    </div>
                  </button>

                  <button 
                    onClick={() => { setActiveTab('reports'); setIsMobileMenuOpen(false); }}
                    className={`group flex items-center justify-between px-3.5 py-2.5 rounded-2xl transition-all duration-200 text-xs ${
                      activeTab === 'reports' 
                        ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white font-extrabold shadow-md shadow-blue-600/20' 
                        : 'text-slate-600 hover:text-blue-700 hover:bg-blue-50/80 font-bold'
                    }`}
                  >
                    <div className="flex items-center gap-3">
                      <BarChart3 size={18} className={activeTab === 'reports' ? 'text-white' : 'text-slate-400 group-hover:text-blue-600 transition-colors'} />
                      <span>RelatÃ³rios</span>
                    </div>
                  </button>

                  {(userProfile?.role === 'LÃDER' || userProfile?.role === 'SETOR') && (
                    <button 
                      onClick={() => { setActiveTab('leader-stats'); setIsMobileMenuOpen(false); }}
                      className={`group flex items-center justify-between px-3.5 py-2.5 rounded-2xl transition-all duration-200 text-xs ${
                        activeTab === 'leader-stats' 
                          ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white font-extrabold shadow-md shadow-blue-600/20' 
                          : 'text-slate-600 hover:text-blue-700 hover:bg-blue-50/80 font-bold'
                      }`}
                    >
                      <div className="flex items-center gap-3">
                        <BarChart3 size={18} className={activeTab === 'leader-stats' ? 'text-white' : 'text-slate-400 group-hover:text-blue-600 transition-colors'} />
                        <span>EstatÃ­sticas</span>
                      </div>
                    </button>
                  )}
                </>
              )}
            </>
          )}
        </nav>

        <div className="mt-auto pt-4 border-t border-blue-100/80 space-y-2">
          <div className="p-3 rounded-2xl bg-slate-50 border border-slate-200/80 hover:border-blue-200 transition-all">
            <p className="text-[10px] font-black text-slate-400 uppercase tracking-wider mb-2">UsuÃ¡rio Conectado</p>
            <div className="flex items-center gap-3">
              <img src={user.photoURL || ''} className="w-9 h-9 rounded-xl border-2 border-blue-500/30 object-cover shadow-sm shrink-0" alt="" />
              <div className="overflow-hidden flex-1 min-w-0">
                <p className="text-xs font-extrabold text-slate-900 truncate">{user.displayName}</p>
                {userProfile?.allowedSectors && userProfile.allowedSectors.length > 1 ? (
                  <select 
                    value={selectedSector}
                    onChange={(e) => setSelectedSector(e.target.value)}
                    className="text-[10px] text-slate-900 font-extrabold uppercase bg-white border border-slate-200 rounded-lg px-2 py-1 mt-1 cursor-pointer hover:border-blue-300 transition-all w-full focus:ring-1 focus:ring-blue-500"
                  >
                    {userProfile.allowedSectors.map(s => (
                      <option key={s} value={s}>{s}</option>
                    ))}
                  </select>
                ) : (
                  <p className="text-[10px] text-slate-500 font-bold truncate uppercase mt-0.5">{selectedSector || 'Sem Setor'}</p>
                )}
                <button 
                  onClick={handleLogout} 
                  className="text-[10px] text-rose-600 font-bold hover:text-rose-700 hover:bg-rose-50 px-2 py-0.5 rounded-md transition-all flex items-center gap-1 mt-1.5"
                >
                  <LogOut size={11} /> Sair do sistema
                </button>
              </div>
            </div>
          </div>

          <button 
            onClick={() => setShowSettingsModal(true)}
            className="flex items-center gap-3 px-3.5 py-2.5 rounded-2xl text-slate-600 hover:text-blue-700 hover:bg-blue-50/80 text-xs font-bold w-full transition-all"
          >
            <Settings size={18} className="text-slate-400" /> ConfiguraÃ§Ãµes
          </button>
        </div>
      </aside>

      {/* Main Content */}
      <main className="lg:ml-64 p-4 lg:p-10 max-w-7xl mx-auto mt-16 lg:mt-0">
        <header className="flex flex-col lg:flex-row lg:justify-between lg:items-center gap-4 mb-6 lg:mb-10">
          <div>
            <h2 className="text-xl lg:text-3xl font-bold tracking-tight mb-1">
              {activeTab === 'dashboard' && 'VisÃ£o Geral'}
              {activeTab === 'inventory' && 'Gerenciamento de Estoque'}
              {activeTab === 'history' && 'HistÃ³rico de MovimentaÃ§Ãµes'}
              {activeTab === 'requests' && 'SolicitaÃ§Ãµes de Materiais'}
              {activeTab === 'admin-devolutions' && 'DevoluÃ§Ãµes de Materiais'}
              {activeTab === 'trash' && 'Lixeira (ExclusÃ£o em 3 dias)'}
              {activeTab === 'my-requests' && `Minhas SolicitaÃ§Ãµes - ${selectedSector || ''}`}
              {activeTab === 'new-request' && `Nova SolicitaÃ§Ã£o - ${selectedSector || ''}`}
              {activeTab === 'devolution' && `DevoluÃ§Ã£o de Materiais - ${selectedSector || ''}`}
              {editingRequest && ' - Editando SolicitaÃ§Ã£o'}
              {activeTab === 'reports' && 'RelatÃ³rios e AnÃ¡lises'}
              {activeTab === 'leader-stats' && 'EstatÃ­sticas do Almoxarifado'}
            </h2>
              {activeTab === 'dashboard' && (
                <div className="flex items-center gap-4 mt-2">
                  <p className="text-[#78716C]">
                    {new Date().toLocaleDateString('pt-BR', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}
                  </p>
                  {isAdmin && (
                    <div className="flex items-center gap-2 bg-white px-3 py-1.5 rounded-xl border border-[#E7E5E4]">
                      <Package size={14} className="text-[#A8A29E]" />
                      <select 
                        className="text-xs font-bold focus:outline-none bg-transparent"
                        value={inventoryLocation}
                        onChange={e => setInventoryLocation(e.target.value as 'Almoxarifado' | 'FarmÃ¡cia')}
                      >
                        <option value="Almoxarifado">Almoxarifado</option>
                        <option value="FarmÃ¡cia">FarmÃ¡cia</option>
                      </select>
                    </div>
                  )}
                </div>
              )}
              {activeTab === 'history' && (
                <p className="text-[#78716C]">
                  {new Date().toLocaleDateString('pt-BR', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}
                </p>
              )}
              {activeTab === 'reports' && (
                <div className="flex flex-wrap items-center gap-3 mt-3">
                  <div className="flex items-center gap-2 bg-white px-3.5 py-2 rounded-2xl border border-slate-200/90 shadow-sm hover:border-blue-300 transition-all">
                    <Calendar size={15} className="text-blue-600" />
                    <input 
                      type="date" 
                      className="text-xs font-extrabold text-slate-700 focus:outline-none cursor-pointer"
                      value={reportRange.start}
                      onChange={e => setReportRange({...reportRange, start: e.target.value})}
                    />
                    <span className="text-slate-400 text-xs font-bold">atÃ©</span>
                    <input 
                      type="date" 
                      className="text-xs font-extrabold text-slate-700 focus:outline-none cursor-pointer"
                      value={reportRange.end}
                      onChange={e => setReportRange({...reportRange, end: e.target.value})}
                    />
                  </div>
                  {isAdmin && (
                    <div className="flex items-center gap-2 bg-white px-3.5 py-2 rounded-2xl border border-slate-200/90 shadow-sm hover:border-blue-300 transition-all">
                      <Filter size={15} className="text-blue-600" />
                      <select 
                        className="text-xs font-extrabold text-slate-700 focus:outline-none bg-transparent cursor-pointer"
                        value={reportSectorFilter}
                        onChange={e => setReportSectorFilter(e.target.value)}
                      >
                        <option value="all">Todos os Setores</option>
                        {SECTORS.map(s => (
                          <option key={s} value={s}>{s}</option>
                        ))}
                      </select>
                    </div>
                  )}
                  {isAdmin && (
                    <button 
                      onClick={handleExportExcel}
                      className="flex items-center gap-2 bg-gradient-to-r from-emerald-600 to-teal-700 text-white px-4 py-2 rounded-2xl text-xs font-extrabold hover:from-emerald-700 hover:to-teal-800 transition-all shadow-md shadow-emerald-600/20"
                    >
                      <Download size={15} /> Exportar Excel
                    </button>
                  )}
                  {!isAdmin && (
                    <button 
                      onClick={handleExportMaterialsCatalogPDF}
                      className="flex items-center gap-2 bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white px-4 py-2 rounded-2xl text-xs font-extrabold hover:from-blue-800 hover:to-indigo-950 transition-all shadow-md shadow-blue-600/20"
                    >
                      <FileText size={15} /> CatÃ¡logo de Itens
                    </button>
                  )}
                </div>
              )}
              {(activeTab === 'requests' || activeTab === 'my-requests' || activeTab === 'admin-devolutions' || activeTab === 'devolution') && (
                <div className="flex items-center gap-4 mt-2">
                  <button 
                    onClick={handleExportRequestsPDF}
                    className="flex items-center gap-2 bg-rose-600 text-white px-4 py-1.5 rounded-xl text-xs font-bold hover:bg-rose-700 transition-all shadow-sm"
                  >
                    <Download size={14} /> Exportar PDF
                  </button>
                </div>
              )}
            </div>
          
          <div className="flex gap-4 items-center">
            <div className="relative">
              <button 
                onClick={() => setShowNotifications(!showNotifications)}
                className="relative p-2 text-[#57534E] hover:bg-white hover:shadow-sm rounded-xl transition-all"
                title="NotificaÃ§Ãµes"
              >
                <Bell size={24} />
                {notifications.filter(n => !n.read).length > 0 && (
                  <span className="absolute top-1 right-1 w-5 h-5 bg-rose-600 text-white text-[10px] font-bold rounded-full flex items-center justify-center border-2 border-[#F5F5F4]">
                    {notifications.filter(n => !n.read).length}
                  </span>
                )}
              </button>

              <AnimatePresence>
                {showNotifications && (
                  <>
                    <div 
                      className="fixed inset-0 z-40" 
                      onClick={() => setShowNotifications(false)} 
                    />
                    <motion.div 
                      initial={{ opacity: 0, y: 10, scale: 0.95 }}
                      animate={{ opacity: 1, y: 0, scale: 1 }}
                      exit={{ opacity: 0, y: 10, scale: 0.95 }}
                      className="absolute right-0 mt-2 w-80 bg-white rounded-3xl shadow-2xl border border-[#E7E5E4] z-50 overflow-hidden"
                    >
                      <div className="p-4 border-b border-[#E7E5E4] flex justify-between items-center bg-[#FAFAF9]">
                        <h3 className="font-black text-sm">NotificaÃ§Ãµes</h3>
                        <button 
                          onClick={async () => {
                            const unreadSystem = notifications.filter(n => !n.read && n.userId !== 'ADMIN_GROUP');
                            for (const n of unreadSystem) {
                              await updateDoc(doc(db, 'notifications', n.id), { read: true });
                            }
                          }}
                          className="text-[10px] font-bold text-blue-600 hover:underline uppercase tracking-wider"
                        >
                          Limpar Lidas
                        </button>
                      </div>
                      <div className="max-h-[400px] overflow-y-auto">
                        {notifications.length === 0 ? (
                          <div className="p-10 text-center">
                            <Bell size={40} className="mx-auto text-[#E7E5E4] mb-3" />
                            <p className="text-xs text-[#A8A29E] font-medium">Nenhuma notificaÃ§Ã£o</p>
                          </div>
                        ) : (
                          <div className="divide-y divide-[#E7E5E4]">
                            {notifications.filter(n => !n.read).map(n => (
                              <div key={n.id} className={`p-4 hover:bg-[#FAFAF9] transition-colors ${n.type === 'STOCK_ZERO' ? 'bg-rose-50/30' : ''}`}>
                                <div className="flex gap-3">
                                  <div className={`shrink-0 w-8 h-8 rounded-full flex items-center justify-center ${
                                    n.type === 'STOCK_ZERO' ? 'bg-rose-100 text-rose-600' : 
                                    n.type === 'REQUEST' ? 'bg-blue-100 text-blue-600' : 'bg-[#F5F5F4] text-[#78716C]'
                                  }`}>
                                    {n.type === 'STOCK_ZERO' ? <AlertTriangle size={14} /> : <Info size={14} />}
                                  </div>
                                  <div className="flex-1 min-w-0">
                                    <p className="text-xs font-bold text-[#1C1917] mb-0.5">{n.title}</p>
                                    <p className="text-[11px] text-[#57534E] leading-relaxed mb-2">{n.message}</p>
                                    <div className="flex items-center justify-between">
                                      <span className="text-[9px] text-[#A8A29E] font-medium">
                                        {new Date(n.date).toLocaleDateString('pt-BR')} {new Date(n.date).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}
                                      </span>
                                      {n.type === 'STOCK_ZERO' ? (
                                        <button 
                                          onClick={() => setShowStockConfirm({ show: true, notificationId: n.id, itemName: n.itemName })}
                                          className="text-[10px] font-bold text-rose-600 bg-rose-100 px-2 py-1 rounded-lg hover:bg-rose-200 transition-all border border-rose-200"
                                        >
                                          Confirmar CiÃªncia
                                        </button>
                                      ) : (
                                        <button 
                                          onClick={() => updateDoc(doc(db, 'notifications', n.id), { read: true })}
                                          className="text-[10px] font-bold text-blue-600 hover:bg-blue-50 px-2 py-1 rounded-lg transition-all"
                                        >
                                          Marcar como lida
                                        </button>
                                      )}
                                    </div>
                                  </div>
                                </div>
                              </div>
                            ))}
                            {notifications.filter(n => n.read).length > 0 && notifications.filter(n => !n.read).length === 0 && (
                               <div className="p-10 text-center">
                                 <CheckCircle size={40} className="mx-auto text-emerald-100 mb-3" />
                                 <p className="text-xs text-[#A8A29E] font-medium">Tudo em dia!</p>
                               </div>
                            )}
                          </div>
                        )}
                      </div>
                    </motion.div>
                  </>
                )}
              </AnimatePresence>
            </div>

            <div className="relative">
              <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" size={17} />
              <input 
                type="text" 
                placeholder="Buscar insumos e lotes..."
                className="pl-10 pr-4 py-2 bg-white border border-slate-200 rounded-2xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 w-64 text-slate-800 shadow-sm transition-all"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
            </div>
            {activeTab === 'inventory' && (
              <div className="flex flex-wrap items-center gap-2">
                <div className="flex items-center gap-2 bg-white px-3 py-2 rounded-2xl border border-slate-200/90 shadow-sm hover:border-blue-300 transition-all">
                  <Filter size={15} className="text-blue-600" />
                  <select 
                    className="text-xs font-extrabold text-slate-700 focus:outline-none bg-transparent cursor-pointer"
                    value={categoryFilter}
                    onChange={e => setCategoryFilter(e.target.value)}
                  >
                    <option value="all">Todos os Tipos</option>
                    {Array.from(new Set([...Object.keys(CATEGORY_COLORS), ...categories, ...items.map(i => i.category).filter(Boolean)])).sort().map(cat => (
                      <option key={cat} value={cat}>{cat}</option>
                    ))}
                  </select>
                </div>
                <div className="flex items-center gap-2 bg-white px-3 py-2 rounded-2xl border border-slate-200/90 shadow-sm hover:border-blue-300 transition-all">
                  <TrendingUp size={15} className="text-blue-600" />
                  <select 
                    className="text-xs font-extrabold text-slate-700 focus:outline-none bg-transparent cursor-pointer"
                    value={inventorySort}
                    onChange={e => setInventorySort(e.target.value as any)}
                  >
                    <option value="name_asc">A-Z (Nome)</option>
                    <option value="name_desc">Z-A (Nome)</option>
                    <option value="duration_asc">DuraÃ§Ã£o (Menor-Maior)</option>
                    <option value="duration_desc">DuraÃ§Ã£o (Maior-Menor)</option>
                  </select>
                </div>
                <div className="flex items-center gap-2 bg-white px-3 py-2 rounded-2xl border border-slate-200/90 shadow-sm hover:border-blue-300 transition-all">
                  <Filter size={15} className="text-blue-600" />
                  <select 
                    className="text-xs font-extrabold text-slate-700 focus:outline-none bg-transparent cursor-pointer"
                    value={originFilter}
                    onChange={e => setOriginFilter(e.target.value as any)}
                  >
                    <option value="all">Todas Origens</option>
                    <option value="contract">Contrato</option>
                    <option value="extra">Extra</option>
                    <option value="donation">DoaÃ§Ã£o</option>
                  </select>
                </div>
                {isAdmin && (
                  <div className="flex gap-2">
                    <button 
                      onClick={() => {
                        setCategoryModalMaterial('');
                        setCategoryModalNewCategory('');
                        setCustomModalCategory('');
                        setShowChangeCategoryModal(true);
                      }}
                      className="px-3 py-2 bg-indigo-50 border border-indigo-200/90 rounded-2xl text-indigo-700 hover:text-indigo-800 hover:border-indigo-300 hover:bg-indigo-100 transition-all shadow-xs flex items-center gap-1.5 text-xs font-bold"
                      title="Alterar Categoria do Material / Insumo"
                    >
                      <Tag size={15} className="text-indigo-600" />
                      <span className="hidden sm:inline">Alterar Categoria</span>
                    </button>
                    <button 
                      onClick={handleExportLowStockPDF}
                      className="px-3 py-2 bg-amber-50 border border-amber-200 rounded-2xl text-amber-700 hover:text-amber-800 hover:border-amber-300 hover:bg-amber-100 transition-all shadow-sm flex items-center gap-1.5 text-xs font-bold"
                      title="Imprimir RelatÃ³rio de Itens CrÃ­ticos / Estoque Baixo"
                    >
                      <Printer size={16} className="text-amber-600" />
                      <span className="hidden sm:inline">RelatÃ³rio Estoque Baixo</span>
                    </button>
                    <button 
                      onClick={handleExportInventory}
                      className="p-2 bg-white border border-slate-200 rounded-2xl text-slate-600 hover:text-blue-700 hover:border-blue-300 hover:bg-blue-50/50 transition-all shadow-sm"
                      title="Baixar Planilha Excel"
                    >
                      <Download size={18} />
                    </button>
                    <button 
                      onClick={handleExportInventoryPDF}
                      className="p-2 bg-white border border-slate-200 rounded-2xl text-rose-600 hover:text-rose-700 hover:border-rose-300 hover:bg-rose-50 transition-all shadow-sm"
                      title="Baixar RelatÃ³rio PDF de Todo Estoque"
                    >
                      <Printer size={18} />
                    </button>
                  </div>
                )}
              </div>
            )}
            {activeTab === 'dashboard' && (isAdmin || selectedSector === 'FarmÃ¡cia') && (
              <>
                <button 
                  onClick={() => setShowAddModal(true)}
                  className="bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white px-4.5 py-2 rounded-2xl text-xs font-extrabold flex items-center gap-2 hover:from-blue-800 hover:to-indigo-950 transition-all shadow-md shadow-blue-600/20"
                >
                  <Plus size={18} /> Nova Entrada
                </button>
                <button 
                  onClick={() => setShowTransactionModal({ show: true, type: 'exit' })}
                  className="bg-gradient-to-r from-rose-600 to-rose-700 text-white px-4.5 py-2 rounded-2xl text-xs font-extrabold flex items-center gap-2 hover:from-rose-700 hover:to-rose-800 transition-all shadow-md shadow-rose-600/20"
                >
                  <ArrowUpRight size={18} /> Nova SaÃ­da
                </button>
              </>
            )}
          </div>
        </header>

        <AnimatePresence mode="wait">
          {activeTab === 'dashboard' && (
            <motion.div 
              key="dashboard"
              initial={{ opacity: 0, y: 12 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -12 }}
              className="space-y-8"
            >
              {/* 4 Primary KPI Stats Cards */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
                {/* Card 1: Volume Total */}
                <div className="bg-white rounded-xl border border-blue-100/80 shadow-xs hover:shadow-sm hover:border-blue-200 transition-all duration-200 overflow-hidden group relative">
                  <div className="h-1 w-full bg-gradient-to-r from-blue-600 to-cyan-500" />
                  <div className="p-3.5">
                    <div className="flex items-center justify-between mb-2">
                      <span className="text-[10px] font-black text-slate-500 uppercase tracking-wider">Volume em Estoque</span>
                      <div className="bg-gradient-to-br from-blue-600 to-blue-700 text-white p-1.5 rounded-lg shadow-xs group-hover:scale-105 transition-transform">
                        <Package size={15} />
                      </div>
                    </div>
                    <h3 className="text-xl font-black text-slate-900 tracking-tight">{totalVolume.toLocaleString('pt-BR')}</h3>
                    <div className="mt-2 flex items-center gap-1.5">
                      <span className="px-2 py-0.5 rounded-md text-[10px] font-bold bg-blue-50 text-blue-700 border border-blue-100 uppercase tracking-wider flex items-center gap-1">
                        <span className="w-1.5 h-1.5 rounded-full bg-blue-600" />
                        {groupedArray.length} tipos de insumos
                      </span>
                    </div>
                  </div>
                </div>

                {/* Card 2: PatrimÃ´nio */}
                {(isAdmin || selectedSector === 'FarmÃ¡cia') && (
                  <div className="bg-white rounded-xl border border-indigo-100/80 shadow-xs hover:shadow-sm hover:border-indigo-200 transition-all duration-200 overflow-hidden group relative">
                    <div className="h-1 w-full bg-gradient-to-r from-indigo-600 to-blue-600" />
                    <div className="p-3.5">
                      <div className="flex items-center justify-between mb-2">
                        <span className="text-[10px] font-black text-slate-500 uppercase tracking-wider">PatrimÃ´nio Investido</span>
                        <div className="bg-gradient-to-br from-indigo-600 to-indigo-700 text-white p-1.5 rounded-lg shadow-xs group-hover:scale-105 transition-transform">
                          <DollarSign size={15} />
                        </div>
                      </div>
                      <h3 className="text-lg font-black text-slate-900 tracking-tight select-all">
                        {new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(totalInventoryValue)}
                      </h3>
                      <div className="mt-2 flex items-center gap-1.5">
                        <span className="px-2 py-0.5 rounded-md text-[10px] font-bold bg-indigo-50 text-indigo-700 border border-indigo-100 uppercase tracking-wider flex items-center gap-1">
                          <span className="w-1.5 h-1.5 rounded-full bg-indigo-600" />
                          Valor financeiro ativo
                        </span>
                      </div>
                    </div>
                  </div>
                )}

                {/* Card 3: PendÃªncias / SolicitaÃ§Ãµes */}
                <div 
                  onClick={() => setActiveTab('requests')}
                  className="bg-white rounded-xl border border-sky-100/80 shadow-xs hover:shadow-sm hover:border-sky-300 transition-all duration-200 overflow-hidden group cursor-pointer relative"
                >
                  <div className="h-1 w-full bg-gradient-to-r from-sky-500 to-blue-600" />
                  <div className="p-3.5">
                    <div className="flex items-center justify-between mb-2">
                      <span className="text-[10px] font-black text-slate-500 uppercase tracking-wider">SolicitaÃ§Ãµes Pendentes</span>
                      <div className="bg-gradient-to-br from-sky-500 to-blue-600 text-white p-1.5 rounded-lg shadow-xs group-hover:scale-105 transition-transform">
                        <Clock size={15} />
                      </div>
                    </div>
                    <h3 className={`text-xl font-black tracking-tight ${pendingRequestsCount > 0 ? 'text-sky-700' : 'text-slate-900'}`}>
                      {pendingRequestsCount}
                    </h3>
                    <div className="mt-2 flex items-center gap-1.5">
                      <span className={`px-2 py-0.5 rounded-md text-[10px] font-bold uppercase tracking-wider flex items-center gap-1 ${pendingRequestsCount > 0 ? 'bg-sky-50 text-sky-800 border border-sky-200' : 'bg-slate-50 text-slate-600 border border-slate-100'}`}>
                        <span className={`w-1.5 h-1.5 rounded-full ${pendingRequestsCount > 0 ? 'bg-sky-500' : 'bg-slate-400'}`} />
                        {pendingRequestsCount > 0 ? 'Aguardando atendimento' : 'Nenhuma pendÃªncia'}
                      </span>
                    </div>
                  </div>
                </div>

                {/* Card 4: NÃ­vel CrÃ­tico / Alertas */}
                <div 
                  onClick={() => totalAlertsCount > 0 && setShowDetailModal({ 
                    show: true, 
                    type: 'all_alerts', 
                    items: [...expiredItems, ...lowStockItems, ...nearExpiryItems] as any 
                  })}
                  className={`bg-white rounded-xl border transition-all duration-200 overflow-hidden group cursor-pointer relative ${
                    totalAlertsCount > 0
                      ? 'border-amber-200/80 shadow-xs hover:border-amber-300 hover:shadow-sm'
                      : 'border-blue-100/80 shadow-xs hover:border-blue-200'
                  }`}
                >
                  <div className={`h-1 w-full ${expiredItems.length > 0 ? 'bg-gradient-to-r from-rose-600 to-amber-500' : lowStockItems.length > 0 ? 'bg-gradient-to-r from-amber-500 to-rose-500' : 'bg-gradient-to-r from-emerald-500 to-blue-500'}`} />
                  <div className="p-3.5">
                    <div className="flex items-center justify-between mb-2">
                      <span className="text-[10px] font-black text-slate-500 uppercase tracking-wider">AtenÃ§Ã£o NecessÃ¡ria</span>
                      <div className={`p-1.5 rounded-lg shadow-xs group-hover:scale-105 transition-transform text-white ${
                        expiredItems.length > 0 ? 'bg-gradient-to-br from-rose-600 to-amber-600' : lowStockItems.length > 0 ? 'bg-gradient-to-br from-amber-500 to-rose-500' : 'bg-gradient-to-br from-emerald-500 to-blue-600'
                      }`}>
                        <AlertTriangle size={15} />
                      </div>
                    </div>
                    <h3 className={`text-xl font-black tracking-tight ${expiredItems.length > 0 ? 'text-rose-600' : lowStockItems.length > 0 ? 'text-amber-600' : 'text-slate-900'}`}>
                      {totalAlertsCount}
                    </h3>
                    <div className="mt-2 flex flex-wrap items-center gap-1">
                      {totalAlertsCount === 0 ? (
                        <span className="px-2 py-0.5 rounded-md text-[10px] font-bold uppercase tracking-wider flex items-center gap-1 bg-emerald-50 text-emerald-800 border border-emerald-200">
                          <span className="w-1.5 h-1.5 rounded-full bg-emerald-500" />
                          Tudo em dia
                        </span>
                      ) : (
                        <>
                          {expiredItems.length > 0 && (
                            <span className="px-1.5 py-0.5 rounded-md text-[9px] font-black uppercase tracking-wider bg-rose-100 text-rose-800 border border-rose-300 flex items-center gap-1">
                              <span className="w-1 h-1 rounded-full bg-rose-600" />
                              {expiredItems.length} vencido{expiredItems.length > 1 ? 's' : ''}
                            </span>
                          )}
                          {lowStockItems.length > 0 && (
                            <span className="px-1.5 py-0.5 rounded-md text-[9px] font-black uppercase tracking-wider bg-amber-100 text-amber-800 border border-amber-300 flex items-center gap-1">
                              <span className="w-1 h-1 rounded-full bg-amber-500" />
                              {lowStockItems.length} baixo estoque
                            </span>
                          )}
                          {nearExpiryItems.length > 0 && (
                            <span className="px-1.5 py-0.5 rounded-md text-[9px] font-black uppercase tracking-wider bg-sky-100 text-sky-800 border border-sky-300 flex items-center gap-1">
                              <span className="w-1 h-1 rounded-full bg-sky-500" />
                              {nearExpiryItems.length} prÃ³x. vencer
                            </span>
                          )}
                        </>
                      )}
                    </div>
                  </div>
                </div>
              </div>

              {/* Category Volume Distribution Bar Chart Section */}
              {categoryDistribution.length > 0 && (
                <div className="bg-white rounded-3xl border border-blue-100 p-6 shadow-sm space-y-4">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-blue-50 pb-4">
                    <div className="flex items-center gap-3">
                      <div className="p-2.5 bg-blue-50 text-blue-700 rounded-2xl border border-blue-100">
                        <BarChart3 size={20} />
                      </div>
                      <div>
                        <h4 className="font-extrabold text-sm text-slate-900 uppercase tracking-wider">
                          DistribuiÃ§Ã£o de Estoque por Categoria
                        </h4>
                        <p className="text-xs text-slate-500">
                          {distribViewMode === 'types' 
                            ? 'Variedade e diversidade por tipo de produto cadastrado (prioridade)' 
                            : 'Volume acumulado por quantidade total de unidades em estoque'}
                        </p>
                      </div>
                    </div>

                    <div className="flex items-center gap-1.5 bg-slate-100 p-1 rounded-2xl border border-slate-200/80 self-start sm:self-auto">
                      <button
                        type="button"
                        onClick={() => setDistribViewMode('types')}
                        className={`px-3 py-1.5 rounded-xl text-xs font-extrabold transition-all ${
                          distribViewMode === 'types'
                            ? 'bg-white text-blue-700 shadow-xs'
                            : 'text-slate-600 hover:text-slate-900'
                        }`}
                      >
                        Por Tipos de Produto
                      </button>
                      <button
                        type="button"
                        onClick={() => setDistribViewMode('units')}
                        className={`px-3 py-1.5 rounded-xl text-xs font-extrabold transition-all ${
                          distribViewMode === 'units'
                            ? 'bg-white text-blue-700 shadow-xs'
                            : 'text-slate-600 hover:text-slate-900'
                        }`}
                      >
                        Por Unidades
                      </button>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 pt-2">
                    {categoryDistribution.map((cat) => {
                      const displayPercentage = distribViewMode === 'types' ? cat.typePercentage : cat.unitPercentage;
                      return (
                        <div 
                          key={cat.category} 
                          onClick={() => {
                            setCategoryFilter(cat.category);
                            setActiveTab('inventory');
                          }}
                          className="p-4 rounded-2xl bg-slate-50/70 border border-slate-100 hover:border-blue-300 hover:bg-blue-50/50 transition-all cursor-pointer group"
                          title={`Clique para ver os produtos da categoria ${cat.category} no estoque`}
                        >
                          <div className="flex items-center justify-between mb-2">
                            <span className="text-xs font-bold text-slate-800 truncate max-w-[170px] group-hover:text-blue-700 transition-colors">{cat.category}</span>
                            <span className="text-xs font-black text-blue-700 bg-blue-100/70 group-hover:bg-blue-600 group-hover:text-white transition-all px-2.5 py-0.5 rounded-md">
                              {distribViewMode === 'types' ? `${cat.count} ${cat.count === 1 ? 'tipo' : 'tipos'}` : `${cat.totalQty.toLocaleString('pt-BR')} un`}
                            </span>
                          </div>
                          <div className="w-full bg-slate-200/80 rounded-full h-2.5 overflow-hidden">
                            <div 
                              className="bg-gradient-to-r from-blue-600 via-indigo-600 to-cyan-500 h-2.5 rounded-full transition-all duration-500"
                              style={{ width: `${displayPercentage}%` }}
                            />
                          </div>
                          <div className="flex justify-between items-center mt-2.5 text-[10px] text-slate-600 font-medium">
                            <span className="font-extrabold text-slate-700">
                              {distribViewMode === 'types' 
                                ? `${cat.totalQty.toLocaleString('pt-BR')} un em estoque` 
                                : `${cat.count} tipos de produtos`}
                            </span>
                            <span>{cat.value > 0 ? new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(cat.value) : ''}</span>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              {/* Bento Grid Split Columns */}
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
                {/* Left Column: Central de Alertas CrÃ­ticos (5/12 cols) */}
                <div className="lg:col-span-5 bg-white rounded-3xl border border-blue-100 shadow-sm overflow-hidden space-y-0">
                  <div className="bg-gradient-to-r from-slate-900 via-blue-950 to-slate-900 text-white p-5 flex items-center justify-between">
                    <div className="flex items-center gap-2.5">
                      <div className="p-2 bg-amber-500/20 text-amber-400 rounded-xl border border-amber-400/30">
                        <AlertTriangle size={18} />
                      </div>
                      <div>
                        <h4 className="text-xs font-extrabold uppercase tracking-wider text-white">
                          Central de Alertas CrÃ­ticos
                        </h4>
                        <p className="text-[10px] text-blue-200 font-medium">
                          Itens vencidos, com estoque baixo ou prÃ³ximos ao vencimento
                        </p>
                      </div>
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={handleExportLowStockPDF}
                        className="px-2.5 py-1 bg-amber-500/20 hover:bg-amber-500/30 text-amber-200 border border-amber-400/40 rounded-xl text-[11px] font-extrabold transition-all flex items-center gap-1 shadow-xs"
                        title="Imprimir RelatÃ³rio PDF dos Itens CrÃ­ticos / Estoque Baixo"
                      >
                        <Printer size={13} /> RelatÃ³rio PDF
                      </button>
                      <span className="px-2.5 py-1 rounded-full text-xs font-black bg-amber-500/20 text-amber-300 border border-amber-400/30">
                        {totalAlertsCount}
                      </span>
                    </div>
                  </div>

                  <div className="p-5 space-y-3 max-h-[480px] overflow-y-auto">
                    {totalAlertsCount === 0 && (
                      <div className="py-14 text-center">
                        <CheckCircle size={40} className="mx-auto text-emerald-500 mb-3" />
                        <p className="text-sm text-slate-800 font-bold mb-1">Estoque 100% em Conformidade</p>
                        <p className="text-xs text-slate-500 max-w-xs mx-auto">
                          Nenhum insumo apresentou nÃ­vel crÃ­tico de reposiÃ§Ã£o, vencimento ultrapassado ou data de expiraÃ§Ã£o prÃ³xima.
                        </p>
                      </div>
                    )}

                    {/* Expired Items - High Urgency Red */}
                    {expiredItems.map(item => (
                      <div key={`expired-${item.id}`} className="flex items-center justify-between p-3.5 bg-rose-50/90 rounded-2xl border border-rose-200 hover:bg-rose-100/80 transition-all duration-200">
                        <div className="flex items-center gap-3 min-w-0 flex-1">
                          <div className="w-9 h-9 rounded-xl flex items-center justify-center shrink-0 border bg-rose-100 text-rose-700 border-rose-300 font-black">
                            <Calendar size={16} />
                          </div>
                          <div className="min-w-0 flex-1">
                            <div className="flex items-center gap-1.5">
                              <p className="font-extrabold text-xs text-slate-900 truncate leading-tight">{item.name}</p>
                              <span className="text-[9px] font-black uppercase px-1.5 py-0.2 rounded bg-rose-200 text-rose-800 border border-rose-300">Vencido</span>
                            </div>
                            <p className="text-[10px] font-bold mt-0.5 text-rose-700">
                              Expirou em: {new Date(item.expiry_date!).toLocaleDateString('pt-BR')} ({item.quantity} un em estoque)
                            </p>
                          </div>
                        </div>
                        <button 
                          onClick={() => setShowTransactionModal({ show: true, type: 'exit', item })}
                          className="px-3 py-1.5 rounded-xl text-xs font-bold bg-rose-600 text-white hover:bg-rose-700 transition-all shadow-sm shrink-0 ml-2"
                        >
                          Retirar
                        </button>
                      </div>
                    ))}

                    {/* Low Stock Items */}
                    {lowStockItems.map(group => (
                      <div key={`low-${group.name}`} className="flex items-center justify-between p-3.5 bg-amber-50/50 rounded-2xl border border-amber-200/60 hover:bg-amber-50 transition-all duration-200">
                        <div className="flex items-center gap-3 min-w-0">
                          <div className="w-9 h-9 bg-amber-100 text-amber-900 rounded-xl flex items-center justify-center font-black text-xs shrink-0 border border-amber-200">
                            {group.total_quantity}
                          </div>
                          <div className="min-w-0 flex-1">
                            <div className="flex items-center gap-1.5">
                              <p className="font-extrabold text-xs text-slate-900 truncate leading-tight">{group.name}</p>
                              <span className="text-[9px] font-black uppercase px-1.5 py-0.2 rounded bg-amber-200 text-amber-900">Estoque Baixo</span>
                            </div>
                            <p className="text-[10px] text-amber-800 font-semibold mt-0.5">
                              Abaixo do mÃ­nimo recomendado ({group.min_quantity} un)
                            </p>
                          </div>
                        </div>
                        <button 
                          onClick={() => setShowTransactionModal({ show: true, type: 'entry', item: group.batches[0] })}
                          className="bg-gradient-to-r from-amber-600 to-amber-700 text-white px-3 py-1.5 rounded-xl text-xs font-bold hover:from-amber-700 hover:to-amber-800 transition-all shadow-sm shrink-0 ml-2"
                        >
                          Repor
                        </button>
                      </div>
                    ))}

                    {/* Near Expiry Items */}
                    {nearExpiryItems.map(item => (
                      <div key={`exp-${item.id}`} className="flex items-center justify-between p-3.5 bg-sky-50/60 rounded-2xl border border-sky-200/80 hover:bg-sky-50 transition-all duration-200">
                        <div className="flex items-center gap-3 min-w-0 flex-1">
                          <div className="w-9 h-9 rounded-xl flex items-center justify-center shrink-0 border bg-sky-100 text-sky-800 border-sky-200">
                            <Calendar size={16} />
                          </div>
                          <div className="min-w-0 flex-1">
                            <div className="flex items-center gap-1.5">
                              <p className="font-extrabold text-xs text-slate-900 truncate leading-tight">{item.name}</p>
                              <span className="text-[9px] font-black uppercase px-1.5 py-0.2 rounded bg-sky-200 text-sky-900">PrÃ³x. Vencer</span>
                            </div>
                            <p className="text-[10px] font-bold mt-0.5 text-sky-800">
                              Vence em: {new Date(item.expiry_date!).toLocaleDateString('pt-BR')} ({item.quantity} un)
                            </p>
                          </div>
                        </div>
                        <button 
                          onClick={() => setShowTransactionModal({ show: true, type: 'exit', item })}
                          className="px-3 py-1.5 rounded-xl text-xs font-bold bg-slate-800 text-white hover:bg-slate-900 transition-all shadow-sm shrink-0 ml-2"
                        >
                          Retirar
                        </button>
                      </div>
                    ))}
                  </div>
                </div>

                {/* Right Column: MovimentaÃ§Ãµes Recentes (7/12 cols) */}
                <div className="lg:col-span-7 bg-white rounded-3xl border border-blue-100 shadow-sm overflow-hidden space-y-0">
                  <div className="bg-gradient-to-r from-slate-900 via-blue-950 to-slate-900 text-white p-5 flex items-center justify-between">
                    <div className="flex items-center gap-2.5">
                      <div className="p-2 bg-blue-500/20 text-blue-300 rounded-xl border border-blue-400/30">
                        <History size={18} />
                      </div>
                      <div>
                        <h4 className="text-xs font-extrabold uppercase tracking-wider text-white">
                          MovimentaÃ§Ãµes Recentes do Estoque
                        </h4>
                        <p className="text-[10px] text-blue-200 font-medium">
                          HistÃ³rico de saÃ­das e entradas registradas no almoxarifado
                        </p>
                      </div>
                    </div>
                    <span className="text-[10px] font-bold uppercase tracking-wider bg-blue-500/20 text-blue-300 border border-blue-400/30 px-2.5 py-1 rounded-full">
                      Ãšltimos 5 registros
                    </span>
                  </div>

                  <div className="p-5 space-y-3 max-h-[480px] overflow-y-auto">
                    {recentTransactions.length === 0 && (
                      <div className="py-14 text-center">
                        <History size={40} className="mx-auto text-slate-300 mb-3" />
                        <p className="text-sm text-slate-800 font-bold">Sem registros no momento</p>
                        <p className="text-xs text-slate-500">Nenhuma movimentaÃ§Ã£o realizada nesta localizaÃ§Ã£o.</p>
                      </div>
                    )}

                    {recentTransactions.map(t => (
                      <div key={t.id} className="group flex gap-3.5 p-3 hover:bg-blue-50/40 rounded-2xl transition-all duration-200 border border-slate-100 hover:border-blue-200">
                        <div className={`w-10 h-10 rounded-xl flex items-center justify-center shrink-0 border ${t.type === 'entry' ? 'bg-emerald-50 text-emerald-700 border-emerald-200' : 'bg-blue-50 text-blue-700 border-blue-200'}`}>
                          {t.type === 'entry' ? <ArrowDownLeft size={18} /> : <ArrowUpRight size={18} />}
                        </div>
                        <div className="flex-1 min-w-0">
                          <div className="flex items-start justify-between gap-2">
                            <p className="font-extrabold text-xs text-slate-900 truncate leading-tight" title={t.item_name}>
                              {t.item_name}
                            </p>
                            <span className={`text-xs font-black shrink-0 px-2 py-0.5 rounded-md ${t.type === 'entry' ? 'bg-emerald-100 text-emerald-800' : 'bg-blue-100 text-blue-800'}`}>
                              {t.type === 'entry' ? '+' : '-'}{t.quantity} un
                            </span>
                          </div>
                          <p className="text-[10px] text-slate-500 font-medium mt-0.5">
                            {t.type === 'entry' ? 'Entrada / AdiÃ§Ã£o em estoque' : `SaÃ­da e entrega p/ setor: ${t.sector || '---'}`}
                          </p>
                          <div className="flex flex-wrap items-center gap-x-2 gap-y-1 mt-2 pt-1.5 border-t border-dashed border-slate-200/80">
                            <span className="text-[10px] text-slate-400 font-medium flex items-center gap-1">
                              <Clock size={11} />
                              {new Date(t.date).toLocaleDateString('pt-BR')} Ã s {new Date(t.date).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}
                            </span>
                            {t.responsible && (
                              <span className="text-[9px] text-blue-800 font-bold bg-blue-50 border border-blue-100 px-2 py-0.5 rounded-md">
                                {t.responsible.split('@')[0]}
                              </span>
                            )}
                            {t.supplier && (
                              <span className="text-[9px] text-amber-800 font-bold bg-amber-50 border border-amber-100 px-2 py-0.5 rounded-md">
                                Forn: {t.supplier}
                              </span>
                            )}
                          </div>
                        </div>
                        {isAdmin && !t.deletedAt && (
                          <button 
                            onClick={() => {
                              setDeletionReason('');
                              setShowDeleteModal({ show: true, transactionId: t.id });
                            }}
                            className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-xl transition-all self-center opacity-0 group-hover:opacity-100 focus:opacity-100 shrink-0"
                            title="Excluir MovimentaÃ§Ã£o"
                          >
                            <Trash2 size={15} />
                          </button>
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'inventory' && (
            <motion.div 
              key="inventory"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="space-y-4"
            >
              <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 bg-white p-4 rounded-3xl border border-blue-100/80 shadow-sm">
                {isAdmin ? (
                  <div className="flex items-center gap-2 bg-slate-100/80 p-1.5 rounded-2xl border border-slate-200/80">
                    <button 
                      onClick={() => setInventoryLocation('Almoxarifado')}
                      className={`px-5 py-2 rounded-xl text-xs font-extrabold transition-all duration-200 flex items-center gap-2 ${
                        inventoryLocation === 'Almoxarifado' 
                          ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white shadow-md shadow-blue-600/20' 
                          : 'text-slate-600 hover:bg-slate-200/70'
                      }`}
                    >
                      <Package size={15} /> Almoxarifado Geral
                    </button>
                    <button 
                      onClick={() => setInventoryLocation('FarmÃ¡cia')}
                      className={`px-5 py-2 rounded-xl text-xs font-extrabold transition-all duration-200 flex items-center gap-2 ${
                        inventoryLocation === 'FarmÃ¡cia' 
                          ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white shadow-md shadow-blue-600/20' 
                          : 'text-slate-600 hover:bg-slate-200/70'
                      }`}
                    >
                      <Users size={15} /> Estoque FarmÃ¡cia
                    </button>
                  </div>
                ) : (
                  <div className="flex items-center gap-3 px-4 py-2 bg-blue-50/80 rounded-2xl border border-blue-100">
                    <div className="p-2 bg-gradient-to-br from-blue-700 to-indigo-900 text-white rounded-xl shadow-sm">
                      {inventoryLocation === 'FarmÃ¡cia' ? <Users size={16} /> : <Package size={16} />}
                    </div>
                    <div>
                      <p className="text-xs font-black text-slate-900">
                        Estoque: <span className="text-blue-700">{inventoryLocation === 'FarmÃ¡cia' ? 'Medicamentos (FarmÃ¡cia)' : 'Almoxarifado Geral'}</span>
                      </p>
                      <p className="text-[10px] text-slate-500 font-medium">Acesso exclusivo aos medicamentos da FarmÃ¡cia</p>
                    </div>
                  </div>
                )}

                <div className="flex items-center gap-3">
                  {inventoryLocation === 'FarmÃ¡cia' && (
                    <button 
                      onClick={() => setActiveTab('new-request')}
                      className="px-4 py-2.5 bg-gradient-to-r from-emerald-600 to-teal-700 hover:from-emerald-700 hover:to-teal-800 text-white font-extrabold text-xs rounded-2xl shadow-md hover:shadow-lg transition-all flex items-center gap-2"
                      title="Solicitar novos medicamentos ao Almoxarifado Geral"
                    >
                      <Plus size={16} /> Solicitar ao Almoxarifado
                    </button>
                  )}
                  <span className="text-xs font-black text-slate-500 uppercase tracking-wider bg-slate-50 border border-slate-200 px-3 py-1.5 rounded-xl">
                    VisualizaÃ§Ã£o: <span className="text-blue-700 font-black">{inventoryLocation}</span>
                  </span>
                </div>
              </div>

              <div className="bg-white rounded-3xl border border-blue-100 shadow-sm overflow-hidden">
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse min-w-[1000px]">
                <thead>
                  <tr className="bg-gradient-to-r from-slate-900 via-blue-950 to-slate-900 text-white">
                    <th className="px-6 py-4 font-black text-xs text-blue-200/90 uppercase tracking-wider">Item / Insumo</th>
                    <th className="px-6 py-4 font-black text-xs text-blue-200/90 uppercase tracking-wider">Categoria {isAdmin && '/ Fornecedor'}</th>
                    {isAdmin && <th className="px-6 py-4 font-black text-xs text-blue-200/90 uppercase tracking-wider">Origem</th>}
                    <th className="px-6 py-4 font-black text-xs text-blue-200/90 uppercase tracking-wider">{isAdmin ? 'PreÃ§o Un.' : '---'}</th>
                    <th className="px-6 py-4 font-black text-xs text-blue-200/90 uppercase tracking-wider text-center">Quantidade</th>
                    <th className="px-6 py-4 font-black text-xs text-blue-200/90 uppercase tracking-wider">MÃ­nimo</th>
                    <th className="px-6 py-4 font-black text-xs text-blue-200/90 uppercase tracking-wider text-center">DuraÃ§Ã£o</th>
                    <th className="px-6 py-4 font-black text-xs text-blue-200/90 uppercase tracking-wider">Status CrÃ­tico</th>
                    <th className="px-6 py-4 font-black text-xs text-blue-200/90 uppercase tracking-wider text-right">AÃ§Ãµes</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-blue-50/80">
                  {groupedArray.map(group => (
                    <React.Fragment key={group.name}>
                      <tr 
                        className="bg-white hover:bg-blue-50/40 transition-all cursor-pointer group/row"
                        onClick={() => toggleExpand(group.name)}
                      >
                        <td className="px-6 py-4.5">
                          <div className="flex items-center gap-3">
                            <div className={`p-1.5 rounded-lg bg-slate-100 text-slate-500 group-hover/row:bg-blue-100 group-hover/row:text-blue-700 transition-all ${expandedItems.has(group.name) ? 'rotate-90 bg-blue-100 text-blue-700' : ''}`}>
                              <ChevronRight size={16} />
                            </div>
                            {isAdmin && editingMaterialName?.oldName === group.name ? (
                              <div className="flex flex-col gap-2 p-3 bg-slate-50 border border-slate-200 rounded-2xl shadow-sm" onClick={(e) => e.stopPropagation()}>
                                <div className="flex items-center gap-2">
                                  <input 
                                    type="text" 
                                    value={editingMaterialName.newName}
                                    onChange={(e) => setEditingMaterialName({ ...editingMaterialName, newName: e.target.value })}
                                    className="px-3 py-1 bg-white border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 font-bold text-sm text-slate-900"
                                    autoFocus
                                  />
                                  <button 
                                    onClick={handleUpdateMaterialName}
                                    className="p-1.5 text-emerald-600 hover:bg-emerald-50 rounded-xl"
                                    title="Salvar"
                                  >
                                    <Check size={18} />
                                  </button>
                                  <button 
                                    onClick={() => setEditingMaterialName(null)}
                                    className="p-1.5 text-rose-600 hover:bg-rose-50 rounded-xl"
                                    title="Cancelar"
                                  >
                                    <X size={18} />
                                  </button>
                                </div>
                                {group.category === 'Medicamentos' && (
                                  <div className="flex flex-col gap-1 bg-white p-2 rounded-xl border border-slate-200">
                                    <div className="flex flex-wrap gap-1 items-center">
                                      <span className="text-[8px] font-black text-slate-500 uppercase tracking-wider mr-1">Unidades:</span>
                                      {['mg', 'mcg', 'UI', 'g', 'ml', '%'].map(unit => (
                                        <button
                                          key={unit}
                                          type="button"
                                          onClick={() => {
                                            let currentName = editingMaterialName.newName.trim();
                                            if (currentName) {
                                              if (!currentName.endsWith(' ')) {
                                                currentName += ' ';
                                              }
                                              currentName += unit;
                                              setEditingMaterialName({ ...editingMaterialName, newName: currentName });
                                            }
                                          }}
                                          className="px-1.5 py-0.5 bg-slate-100 hover:bg-blue-700 hover:text-white text-slate-700 rounded text-[9px] font-bold transition-all uppercase"
                                        >
                                          +{unit}
                                        </button>
                                      ))}
                                    </div>
                                    <div className="flex flex-wrap gap-1 items-center">
                                      <span className="text-[8px] font-black text-slate-500 uppercase tracking-wider mr-1">Dosagem:</span>
                                      {['500 mg', '1000 mg', '1000 UI', '5000 UI', '10.000 UI', '50.000 UI'].map(dose => (
                                        <button
                                          key={dose}
                                          type="button"
                                          onClick={() => {
                                            let currentName = editingMaterialName.newName.trim();
                                            if (currentName) {
                                              if (!currentName.endsWith(' ')) {
                                                currentName += ' ';
                                              }
                                              currentName += dose;
                                              setEditingMaterialName({ ...editingMaterialName, newName: currentName });
                                            }
                                          }}
                                          className="px-1.5 py-0.5 bg-emerald-50 hover:bg-emerald-600 hover:text-white text-emerald-700 rounded text-[9px] font-bold transition-all uppercase"
                                        >
                                          +{dose}
                                        </button>
                                      ))}
                                    </div>
                                  </div>
                                )}
                              </div>
                            ) : (
                              <div className="flex flex-col">
                                <div className="flex items-center gap-2 group/name flex-wrap">
                                  <p className="font-extrabold text-sm text-slate-900 group-hover/row:text-blue-700 transition-colors">{group.name}</p>
                                  {group.unit_measure && (
                                    <span className="text-[9px] font-black px-2 py-0.5 rounded-md bg-amber-50 text-amber-800 border border-amber-200/80 uppercase tracking-wider">
                                      {group.unit_measure}
                                    </span>
                                  )}
                                  {Array.from(new Set(group.batches.map(b => b.medication_type).filter(Boolean))).map(type => (
                                    <span key={type} className="text-[9px] font-black px-2 py-0.5 rounded-md bg-blue-50 text-blue-700 border border-blue-100/80 uppercase tracking-wider">
                                      {type}
                                    </span>
                                  ))}
                                  {isAdmin && (
                                    <button 
                                      onClick={(e) => { e.stopPropagation(); setEditingMaterialName({ oldName: group.name, newName: group.name }); }}
                                      className="opacity-0 group-hover/name:opacity-100 p-1 text-slate-400 hover:text-blue-700 transition-all"
                                      title="Editar Nome do Material"
                                    >
                                      <Edit2 size={14} />
                                    </button>
                                  )}
                                </div>
                                {group.batches[0]?.description && (
                                  <p className="text-[10px] text-slate-400 italic mt-0.5 line-clamp-1">{group.batches[0].description}</p>
                                )}
                              </div>
                            )}
                          </div>
                        </td>
                        <td className="px-6 py-4.5" onClick={(e) => e.stopPropagation()}>
                          {isAdmin && editingCategory?.name === group.name && !editingCategory?.itemId ? (
                            <div className="flex flex-col gap-1.5 bg-indigo-50/90 p-2.5 border border-indigo-200 rounded-2xl shadow-md min-w-[210px]">
                              <label className="text-[10px] font-black text-indigo-900 uppercase tracking-wider">Nova Categoria:</label>
                              <select 
                                value={editingCategory.currentCategory === '__NEW__' ? '__NEW__' : editingCategory.currentCategory}
                                onChange={(e) => {
                                  const val = e.target.value;
                                  if (val === '__NEW__') {
                                    setEditingCategory({ ...editingCategory, currentCategory: '__NEW__' });
                                  } else {
                                    setEditingCategory({ ...editingCategory, currentCategory: val });
                                  }
                                }}
                                className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-xl text-xs font-extrabold text-slate-800 focus:ring-2 focus:ring-indigo-500"
                              >
                                {categories.map(cat => (
                                  <option key={`cat-select-${cat}`} value={cat}>{cat}</option>
                                ))}
                                <option value="__NEW__">+ Cadastrar Nova Categoria...</option>
                              </select>

                              {editingCategory.currentCategory === '__NEW__' && (
                                <input 
                                  type="text"
                                  placeholder="Digite a nova categoria"
                                  value={customNewCategory}
                                  onChange={(e) => setCustomNewCategory(e.target.value)}
                                  className="w-full px-2.5 py-1 bg-white border border-indigo-300 rounded-xl text-xs font-extrabold text-indigo-900 focus:ring-2 focus:ring-indigo-500"
                                  autoFocus
                                />
                              )}

                              <div className="flex items-center gap-1.5 mt-1">
                                <button 
                                  onClick={() => handleUpdateCategory()}
                                  className="flex-1 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-bold text-xs transition-all flex items-center justify-center gap-1 shadow-xs"
                                  title="Salvar Categoria"
                                >
                                  <Check size={14} /> Salvar
                                </button>
                                <button 
                                  onClick={() => { setEditingCategory(null); setCustomNewCategory(''); }}
                                  className="py-1.5 px-2.5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-xl font-bold text-xs transition-all"
                                  title="Cancelar"
                                >
                                  <X size={14} />
                                </button>
                              </div>
                            </div>
                          ) : (
                            <div className="flex items-center gap-1.5 group/cat">
                              <div>
                                <p className="text-xs font-bold text-slate-800">{group.category || '---'}</p>
                                {isAdmin && <p className="text-[10px] font-medium text-slate-400 mt-0.5">{group.supplier || '---'}</p>}
                              </div>
                              {isAdmin && (
                                <button 
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    setEditingCategory({ name: group.name, currentCategory: group.category || categories[0] || 'Expediente' });
                                    setCustomNewCategory('');
                                  }}
                                  className="opacity-0 group-hover/cat:opacity-100 p-1 text-slate-400 hover:text-indigo-700 hover:bg-indigo-50 rounded-lg transition-all"
                                  title="Alterar Categoria deste Material"
                                >
                                  <Edit2 size={13} />
                                </button>
                              )}
                            </div>
                          )}
                        </td>
                        <td className="px-6 py-4.5">
                          {isAdmin ? (
                            (() => {
                              const origins = new Set(group.batches.map(b => b.origin));
                              if (origins.size === 1) {
                                const origin = Array.from(origins)[0];
                                return (
                                  <span className={`text-[10px] font-extrabold px-2.5 py-1 rounded-md border ${
                                    origin === 'contract' ? 'bg-blue-50 text-blue-700 border-blue-200/80' : 
                                    origin === 'donation' ? 'bg-emerald-50 text-emerald-700 border-emerald-200/80' : 
                                    'bg-indigo-50 text-indigo-700 border-indigo-200/80'
                                  }`}>
                                    {origin === 'contract' ? 'Contrato' : origin === 'donation' ? 'DoaÃ§Ã£o' : 'Extra'}
                                  </span>
                                );
                              }
                              return (
                                <span className="text-[10px] font-extrabold px-2.5 py-1 rounded-md bg-slate-100 text-slate-600 border border-slate-200 uppercase">
                                  {group.batches.length} Lotes
                                </span>
                              );
                            })()
                          ) : (
                            <span className="text-xs text-slate-300">---</span>
                          )}
                        </td>
                        <td className="px-6 py-4.5 font-semibold text-slate-600 text-xs">---</td>
                        <td className="px-6 py-4.5">
                          <div className="flex flex-col items-center justify-center bg-slate-50/90 rounded-2xl py-1.5 px-3 border border-slate-200/80 min-w-[80px]">
                            <span className={`text-base font-black ${group.total_quantity <= (group.min_quantity || 0) ? 'text-amber-600' : 'text-slate-900'}`}>
                              {group.total_quantity}
                            </span>
                            <span className="text-[8px] font-black text-slate-400 uppercase tracking-wider">Estoque Total</span>
                          </div>
                        </td>
                        <td className="px-6 py-4.5 text-xs font-semibold text-slate-600">
                          <div className="flex flex-col">
                            <span className="flex items-center gap-1 font-bold text-slate-800">
                              {group.min_quantity !== undefined && !isNaN(group.min_quantity) ? group.min_quantity : '---'}
                              <TrendingUp size={12} className="text-blue-600" />
                            </span>
                            {group.weeklyExitRate > 0 && <span className="text-[10px] text-slate-400">({group.weeklyExitRate.toFixed(1)}/sem)</span>}
                          </div>
                        </td>
                        <td className="px-6 py-4.5">
                          <div className={`flex flex-col items-center justify-center py-1.5 px-2.5 rounded-xl border ${
                            group.durationWeeks === 'infinite' ? 'bg-blue-50 border-blue-200 text-blue-700' :
                            group.durationWeeks <= 4 ? 'bg-rose-50 border-rose-200 text-rose-700' :
                            group.durationWeeks <= 8 ? 'bg-amber-50 border-amber-200 text-amber-700' :
                            'bg-emerald-50 border-emerald-200 text-emerald-700'
                          }`}>
                            <span className="text-xs font-black">
                              {group.durationWeeks === 'infinite' ? 'Infinito' : `${group.durationWeeks.toFixed(1)} sem`}
                            </span>
                          </div>
                        </td>
                        <td className="px-6 py-4.5 text-xs">
                          {group.durationWeeks !== 'infinite' ? (
                            <span className={`font-black uppercase tracking-tight text-[10px] px-2 py-0.5 rounded-md border ${
                              group.durationWeeks <= 4 ? 'bg-rose-50 text-rose-700 border-rose-200' :
                              group.durationWeeks <= 8 ? 'bg-amber-50 text-amber-700 border-amber-200' :
                              'bg-emerald-50 text-emerald-700 border-emerald-200'
                            }`}>
                              {group.durationWeeks <= 4 ? 'Muito CrÃ­tico' :
                               group.durationWeeks <= 8 ? 'AtenÃ§Ã£o' :
                               'Normal'}
                            </span>
                          ) : (
                            <span className="text-slate-300">---</span>
                          )}
                        </td>
                        <td className="px-6 py-4.5 text-right">
                          <div className="flex flex-col items-end gap-0.5">
                            <button className="text-xs font-extrabold text-blue-700 group-hover/row:text-blue-900 uppercase tracking-wider flex items-center gap-1">
                              {expandedItems.has(group.name) ? 'Recolher' : 'Ver Lotes'}
                            </button>
                            <span className="text-[10px] text-slate-400 font-medium">
                              {group.batches.length} remessas
                            </span>
                          </div>
                        </td>
                      </tr>
                      
                      {expandedItems.has(group.name) && group.batches.map(item => (
                        <tr key={item.id} className="bg-slate-50/70 hover:bg-blue-50/50 transition-all border-l-4 border-blue-600">
                          <td className="px-12 py-3.5">
                            <div className="flex items-center gap-2">
                              <p className="text-xs font-mono font-bold text-slate-800">Lote: {item.batch_number || '---'}</p>
                              {item.unit_measure && (
                                <span className="text-[9px] font-black px-1.5 py-0.5 rounded bg-amber-100/80 text-amber-900 border border-amber-200 uppercase tracking-wide">
                                  {item.unit_measure}
                                </span>
                              )}
                              {item.medication_type && (
                                <span className="text-[9px] font-black px-1.5 py-0.5 rounded bg-slate-200/80 text-slate-700 border border-slate-300/60 uppercase tracking-wide">
                                  {item.medication_type}
                                </span>
                              )}
                            </div>
                            {item.description && <p className="text-[10px] text-slate-400 italic mt-0.5">{item.description}</p>}
                          </td>
                          <td className="px-6 py-3.5">
                            {isAdmin ? (
                              <p className="text-xs text-slate-600 font-medium">{item.supplier || '---'}</p>
                            ) : (
                              <p className="text-xs text-slate-400">---</p>
                            )}
                          </td>
                          <td className="px-6 py-3.5">
                            {isAdmin ? (
                              <span className={`text-[10px] font-bold px-2 py-0.5 rounded-md ${item.origin === 'contract' ? 'bg-blue-100 text-blue-800' : item.origin === 'donation' ? 'bg-emerald-100 text-emerald-800' : 'bg-indigo-100 text-indigo-800'}`}>
                                {item.origin === 'contract' ? 'Contrato' : item.origin === 'donation' ? 'DoaÃ§Ã£o' : 'Extra'}
                              </span>
                            ) : (
                              <span className="text-xs text-slate-300">---</span>
                            )}
                          </td>
                          <td className="px-6 py-3.5 text-xs text-slate-700 font-medium">
                            {isAdmin ? (
                              editingPrice?.id === item.id ? (
                                <div className="flex items-center gap-2" onClick={(e) => e.stopPropagation()}>
                                  <input 
                                    type="number" 
                                    step="0.01"
                                    value={editingPrice.price}
                                    onChange={(e) => setEditingPrice({ ...editingPrice, price: parseFloat(e.target.value) || 0 })}
                                    className="w-24 px-2 py-1 bg-white border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 font-bold text-xs"
                                    autoFocus
                                  />
                                  <button 
                                    onClick={handleUpdatePrice}
                                    className="p-1 text-emerald-600 hover:bg-emerald-50 rounded-md"
                                    title="Salvar"
                                  >
                                    <Check size={14} />
                                  </button>
                                  <button 
                                    onClick={() => setEditingPrice(null)}
                                    className="p-1 text-rose-600 hover:bg-rose-50 rounded-md"
                                    title="Cancelar"
                                  >
                                    <X size={14} />
                                  </button>
                                </div>
                              ) : (
                                <div className="flex items-center gap-2 group/price">
                                  <span className="font-bold text-slate-900">{new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(item.unit_price)}</span>
                                  <button 
                                    onClick={(e) => { e.stopPropagation(); setEditingPrice({ id: item.id, price: item.unit_price }); }}
                                    className="opacity-0 group-hover/price:opacity-100 p-1 text-slate-400 hover:text-blue-700 transition-all"
                                    title="Editar PreÃ§o"
                                  >
                                    <Edit2 size={12} />
                                  </button>
                                </div>
                              )
                            ) : (
                              '---'
                            )}
                          </td>
                          <td className="px-6 py-3.5 text-center">
                            {isAdmin ? (
                              editingQuantity?.id === item.id ? (
                                <div className="flex items-center justify-center gap-2" onClick={(e) => e.stopPropagation()}>
                                  <input 
                                    type="number" 
                                    min="0"
                                    value={editingQuantity.quantity}
                                    onChange={(e) => setEditingQuantity({ ...editingQuantity, quantity: parseInt(e.target.value) || 0 })}
                                    className="w-20 px-2 py-1 bg-white border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 font-bold text-xs"
                                    autoFocus
                                  />
                                  <button 
                                    onClick={handleUpdateQuantity}
                                    className="p-1 text-emerald-600 hover:bg-emerald-50 rounded-md"
                                    title="Salvar"
                                  >
                                    <Check size={14} />
                                  </button>
                                  <button 
                                    onClick={() => setEditingQuantity(null)}
                                    className="p-1 text-rose-600 hover:bg-rose-50 rounded-md"
                                    title="Cancelar"
                                  >
                                    <X size={14} />
                                  </button>
                                </div>
                              ) : (
                                <div className="flex flex-col items-center group/qty">
                                  <div className="flex items-center gap-1.5">
                                    <span className={`text-sm font-black ${item.quantity <= (item.min_quantity || 0) ? 'text-amber-600' : 'text-slate-900'}`}>
                                      {item.quantity} un
                                    </span>
                                    <button 
                                      onClick={(e) => { e.stopPropagation(); setEditingQuantity({ id: item.id, quantity: item.quantity }); }}
                                      className="opacity-0 group-hover/qty:opacity-100 p-1 text-slate-400 hover:text-blue-700 transition-all"
                                      title="Editar Quantidade"
                                    >
                                      <Edit2 size={12} />
                                    </button>
                                  </div>
                                </div>
                              )
                            ) : (
                              <span className={`text-sm font-black ${item.quantity <= (item.min_quantity || 0) ? 'text-amber-600' : 'text-slate-900'}`}>
                                {item.quantity} un
                              </span>
                            )}
                          </td>
                          <td className="px-6 py-3.5 text-xs text-slate-300">---</td>
                          <td className="px-6 py-3.5 text-center">
                            {item.expiry_date ? (
                              <span className={`text-xs font-bold ${item.expiry_date === 'Indeterminada' ? 'text-blue-700' : isNearExpiry(item) ? 'text-rose-600 font-black' : 'text-slate-700'}`}>
                                {item.expiry_date === 'Indeterminada' ? 'Indeterminada' : new Date(item.expiry_date).toLocaleDateString('pt-BR')}
                              </span>
                            ) : (
                              <span className="text-slate-400 text-xs italic">N/A</span>
                            )}
                          </td>
                          <td className="px-6 py-3.5 text-xs text-slate-300">---</td>
                          <td className="px-6 py-3.5 text-right space-x-1.5">
                            <button 
                              onClick={(e) => { e.stopPropagation(); setShowTransactionModal({ show: true, type: 'entry', item }); }}
                              className="p-1.5 text-emerald-700 bg-emerald-50 hover:bg-emerald-100 rounded-lg transition-all border border-emerald-200/60"
                              title="Adicionar Entrada"
                            >
                              <Plus size={15} />
                            </button>
                            <button 
                              onClick={(e) => { e.stopPropagation(); setShowTransactionModal({ show: true, type: 'exit', item }); }}
                              className="p-1.5 text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-lg transition-all border border-blue-200/60"
                              title="Registrar SaÃ­da"
                            >
                              <ArrowUpRight size={15} />
                            </button>
                            {isAdmin && (
                              <button 
                                onClick={(e) => { e.stopPropagation(); handleDeleteItem(item.id); }}
                                className="p-1.5 text-rose-600 bg-rose-50 hover:bg-rose-100 rounded-lg transition-all border border-rose-200/60"
                                title="Excluir Lote"
                              >
                                <Trash2 size={15} />
                              </button>
                            )}
                          </td>
                        </tr>
                      ))}
                    </React.Fragment>
                  ))}
                </tbody>
                <tfoot className="bg-slate-900 text-white border-t-2 border-slate-800">
                  <tr>
                    <td colSpan={4} className="px-6 py-4 font-black text-slate-300 text-right uppercase tracking-wider text-xs">Volume Total em Estoque</td>
                    <td className="px-6 py-4">
                      <div className="flex flex-col items-center justify-center bg-gradient-to-r from-blue-600 to-indigo-700 text-white rounded-2xl py-2 px-3 shadow-md border border-blue-500/30">
                        <span className="text-xl font-black">{totalVolume.toLocaleString('pt-BR')}</span>
                        <span className="text-[8px] font-black uppercase tracking-widest opacity-90">Unidades</span>
                      </div>
                    </td>
                    <td colSpan={4}></td>
                  </tr>
                </tfoot>
              </table>
              </div>
              {filteredItems.length === 0 && (
                <div className="p-16 text-center">
                  <div className="w-16 h-16 bg-blue-50 text-blue-600 rounded-2xl flex items-center justify-center mx-auto mb-3 border border-blue-100">
                    <Package size={32} />
                  </div>
                  <p className="text-slate-900 font-extrabold text-base">Nenhum item encontrado no estoque</p>
                  <p className="text-slate-500 text-xs max-w-sm mx-auto mt-1">Tente ajustar os termos de busca ou selecionar outra categoria nos filtros acima.</p>
                </div>
              )}
              </div>
            </motion.div>
          )}

          {activeTab === 'history' && (
            <motion.div 
              key="history"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="space-y-4"
            >
              <div className="flex justify-between items-center bg-white p-4 rounded-3xl border border-[#E7E5E4] shadow-sm">
                <div className="flex items-center gap-4">
                  <h3 className="text-lg font-bold text-[#1C1917]">HistÃ³rico de MovimentaÃ§Ãµes</h3>
                  {isAdmin && (
                    <div className="flex items-center gap-2 bg-[#F5F5F4] p-1 rounded-2xl border border-[#E7E5E4]">
                      <button 
                        onClick={() => setInventoryLocation('Almoxarifado')}
                        className={`px-4 py-1.5 rounded-xl text-[10px] font-bold transition-all ${inventoryLocation === 'Almoxarifado' ? 'bg-[#1C1917] text-white shadow-sm' : 'text-[#78716C] hover:bg-[#E7E5E4]'}`}
                      >
                        Almoxarifado
                      </button>
                      <button 
                        onClick={() => setInventoryLocation('FarmÃ¡cia')}
                        className={`px-4 py-1.5 rounded-xl text-[10px] font-bold transition-all ${inventoryLocation === 'FarmÃ¡cia' ? 'bg-[#1C1917] text-white shadow-sm' : 'text-[#78716C] hover:bg-[#E7E5E4]'}`}
                      >
                        FarmÃ¡cia
                      </button>
                    </div>
                  )}
                </div>
                <div className="flex gap-2">
                  {showDeletedHistory && transactions.filter(t => !!t.deletedAt).length > 0 && (
                    <button 
                      onClick={handleRecoverAllTransactions}
                      className="px-4 py-2 bg-emerald-600 text-white rounded-xl text-xs font-bold hover:bg-emerald-700 transition-all flex items-center gap-2 shadow-sm"
                    >
                      <RotateCcw size={14} /> Restaurar Tudo
                    </button>
                  )}
                  <button 
                    onClick={() => setShowDeletedHistory(!showDeletedHistory)}
                    className={`px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-2 ${showDeletedHistory ? 'bg-rose-100 text-rose-700' : 'bg-[#F5F5F4] text-[#78716C] hover:bg-[#E7E5E4]'}`}
                  >
                    {showDeletedHistory ? <History size={14} /> : <Trash2 size={14} />}
                    {showDeletedHistory ? 'Ver HistÃ³rico Ativo' : 'Ver ExcluÃ­dos (Testes)'}
                  </button>
                </div>
              </div>

              <div className="bg-white rounded-3xl border border-[#E7E5E4] shadow-sm overflow-hidden">
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse min-w-[1200px]">
                  <thead>
                    <tr className="bg-[#FAFAF9] border-bottom border-[#E7E5E4]">
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Data</th>
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">MovimentaÃ§Ã£o</th>
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Item</th>
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Lote</th>
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Validade</th>
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-center">Origem</th>
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Setor</th>
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">ResponsÃ¡vel</th>
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">Qtd</th>
                    {isAdmin && <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right whitespace-nowrap">Val. Unit</th>}
                    {isAdmin && <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right whitespace-nowrap">Total</th>}
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">AÃ§Ãµes</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#E7E5E4]">
                    {transactions
                      .filter(t => (showDeletedHistory ? !!t.deletedAt : !t.deletedAt) && (t.location || 'Almoxarifado') === inventoryLocation)
                      .map(t => (
                      <tr key={t.id} className={`hover:bg-[#FAFAF9] transition-all ${t.deletedAt ? 'opacity-60 grayscale-[0.5]' : ''}`}>
                        <td className="px-6 py-5 text-sm text-[#57534E] whitespace-nowrap">
                          {new Date(t.date).toLocaleString('pt-BR')}
                        </td>
                        <td className="px-6 py-5">
                          <span className={`inline-flex items-center gap-1 px-3 py-1 rounded-full text-xs font-bold uppercase tracking-wider ${t.type === 'entry' ? 'bg-emerald-100 text-emerald-700' : 'bg-rose-100 text-rose-700'}`}>
                            {t.type === 'entry' ? <ArrowDownLeft size={14} /> : <ArrowUpRight size={14} />}
                            {t.type === 'entry' ? 'Entrada' : 'SaÃ­da'}
                          </span>
                        </td>
                        <td className="px-6 py-5">
                          <div className="font-bold whitespace-nowrap">{t.item_name}</div>
                          {t.exitReason && t.exitReason !== 'consumo' && (
                            <div className="text-[10px] text-rose-500 font-bold mt-1 uppercase">
                              Motivo: {t.exitReason === 'vencido' ? 'Vencimento' : t.exitReason === 'doacao' ? 'DoaÃ§Ã£o' : t.exitReason === 'perda' ? 'Perda/Avaria' : t.exitReason}
                              {t.expiryReason && <span className="text-[#78716C] lowercase font-normal ml-1">({t.expiryReason})</span>}
                            </div>
                          )}
                          {t.deletionReason && (
                            <div className="text-[10px] text-rose-500 font-bold mt-1">ExclusÃ£o: {t.deletionReason}</div>
                          )}
                          {t.deletedByEmail && (
                            <div className="text-[10px] text-rose-400 mt-0.5 italic whitespace-nowrap">Por: {t.deletedByEmail}</div>
                          )}
                        </td>
                        <td className="px-6 py-5 text-xs font-mono text-[#78716C] whitespace-nowrap">
                          {t.batch_number || '---'}
                        </td>
                        <td className="px-6 py-5 text-xs text-[#78716C] whitespace-nowrap">
                          {t.expiry_date ? new Date(t.expiry_date).toLocaleDateString('pt-BR') : '---'}
                        </td>
                        <td className="px-6 py-5 text-center">
                          <span className={`text-[10px] font-bold px-2 py-1 rounded-md ${t.origin === 'contract' ? 'bg-blue-50 text-blue-600' : t.origin === 'donation' ? 'bg-emerald-50 text-emerald-600' : 'bg-purple-50 text-purple-600'}`}>
                            {t.origin === 'contract' ? 'Contrato' : t.origin === 'donation' ? 'DoaÃ§Ã£o' : 'Extra'}
                          </span>
                        </td>
                        <td className="px-6 py-5 text-sm font-medium text-[#78716C]">
                          {t.sector || '---'}
                        </td>
                        <td className="px-6 py-5 text-sm text-[#78716C]">
                          <div className="font-medium">{t.responsible || '---'}</div>
                          <div className="text-[10px] opacity-70">{t.responsibleEmail}</div>
                        </td>
                        <td className="px-6 py-5 text-right font-bold text-lg">
                          {t.quantity}
                        </td>
                        {isAdmin && (
                          <td className="px-6 py-5 text-right text-xs font-medium text-[#78716C]">
                            {(() => {
                              const item = items.find(i => i.id === t.item_id);
                              const price = Number(item?.unit_price) || 0;
                              return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(price);
                            })()}
                          </td>
                        )}
                        {isAdmin && (
                          <td className="px-6 py-5 text-right text-sm font-black text-[#1C1917]">
                            {(() => {
                              const item = items.find(i => i.id === t.item_id);
                              const price = Number(item?.unit_price) || 0;
                              return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(t.quantity * price);
                            })()}
                          </td>
                        )}
                        <td className="px-6 py-5 text-right">
                          <div className="flex items-center justify-end gap-2">
                            {t.type === 'exit' && !t.deletedAt && (
                              <button 
                                onClick={() => {
                                  if (t.exitReason === 'doacao') {
                                    handleExportDonationTermPDF({
                                      donatingUnitName: t.donationUnitName,
                                      receivingUnit: {
                                        name: t.sector || 'Unidade Receptora',
                                        address: t.donationUnitAddress || '',
                                        cnpj: t.donationUnitCNPJ || ''
                                      },
                                      items: [{ product_name: t.item_name, quantity: t.quantity, batch_number: t.batch_number || '---' }],
                                      revisionDate: t.donationRevisionDate || '',
                                      donationNumber: t.donationNumber,
                                      date: t.date
                                    });
                                  } else {
                                    handleExportDeliveryReceiptPDF({
                                      sector: t.sector || 'Sem Setor',
                                      items: [{ product_name: t.item_name, quantity: t.quantity, batch_number: t.batch_number || '---' }],
                                      date: t.date
                                    });
                                  }
                                }}
                                className="p-2 text-blue-600 hover:bg-blue-50 rounded-lg transition-all"
                                title={t.exitReason === 'doacao' ? 'Reimprimir Termo de DoaÃ§Ã£o' : 'Reimprimir Recibo de Entrega'}
                              >
                                {t.exitReason === 'doacao' ? <FileText size={18} /> : <Printer size={18} />}
                              </button>
                            )}
                            {t.deletedAt ? (
                              <button 
                                onClick={() => handleRecoverTransaction(t.id)}
                                className="p-2 text-emerald-600 hover:bg-emerald-50 rounded-lg transition-all"
                                title="Recuperar MovimentaÃ§Ã£o"
                              >
                                <RotateCcw size={18} />
                              </button>
                            ) : (
                              <button 
                                onClick={() => {
                                  setDeletionReason('');
                                  setShowDeleteModal({ show: true, transactionId: t.id });
                                }}
                                className="p-2 text-rose-500 hover:bg-rose-50 rounded-lg transition-all"
                                title="Apagar MovimentaÃ§Ã£o"
                              >
                                <Trash2 size={20} />
                              </button>
                            )}
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                </div>
                {((showDeletedHistory && transactions.filter(t => !!t.deletedAt && (t.location || 'Almoxarifado') === inventoryLocation).length === 0) || 
                  (!showDeletedHistory && transactions.filter(t => !t.deletedAt && (t.location || 'Almoxarifado') === inventoryLocation).length === 0)) && (
                  <div className="p-20 text-center">
                    <History className="mx-auto text-[#E7E5E4] mb-4" size={48} />
                    <p className="text-[#78716C]">Nenhuma movimentaÃ§Ã£o encontrada para {inventoryLocation}.</p>
                  </div>
                )}
              </div>
            </motion.div>
          )}
          {activeTab === 'reports' && (
            <motion.div 
              key="reports"
              initial={{ opacity: 0, scale: 0.98 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.98 }}
              className="space-y-8"
            >
              {/* Executive Reports Banner - Minimalist & Clean Light Theme */}
              <div className="bg-white p-6 sm:p-7 rounded-2xl border border-slate-200/80 shadow-xs text-slate-900">
                <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-6">
                  <div className="space-y-1.5 max-w-2xl">
                    <div className="flex flex-wrap items-center gap-2">
                      <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-extrabold bg-blue-50 text-blue-700 border border-blue-200/80">
                        <BarChart3 size={13} className="text-blue-600" />
                        InteligÃªncia AnalÃ­tica de Estoque
                      </span>
                      <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-slate-100 text-slate-600 border border-slate-200/80">
                        {reportSectorFilter === 'all' ? 'Todos os Setores' : reportSectorFilter}
                      </span>
                    </div>

                    <h2 className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">
                      RelatÃ³rios & GestÃ£o de Consumo
                    </h2>
                    <p className="text-xs sm:text-sm text-slate-500 font-medium leading-relaxed">
                      ProjeÃ§Ãµes orÃ§amentÃ¡rias, histÃ³rico de saÃ­das, curva de movimentaÃ§Ã£o fÃ­sica e relatÃ³rios fiscais do almoxarifado.
                    </p>
                  </div>

                  {/* Minimalist Summary Badges */}
                  <div className="flex items-center gap-3 sm:gap-4 bg-slate-50 p-2.5 sm:p-3 rounded-2xl border border-slate-200/80 shrink-0">
                    <div className="px-3 py-1 text-center">
                      <p className="text-[10px] uppercase font-extrabold text-emerald-600 tracking-wider">Entradas</p>
                      <p className="text-lg font-black text-slate-900 mt-0.5">{reportData.entries}</p>
                    </div>
                    <div className="h-8 w-px bg-slate-200" />
                    <div className="px-3 py-1 text-center">
                      <p className="text-[10px] uppercase font-extrabold text-rose-600 tracking-wider">SaÃ­das</p>
                      <p className="text-lg font-black text-slate-900 mt-0.5">{reportData.exits}</p>
                    </div>
                    <div className="h-8 w-px bg-slate-200" />
                    <div className="px-3 py-1 text-center">
                      <p className="text-[10px] uppercase font-extrabold text-amber-600 tracking-wider">DevoluÃ§Ãµes</p>
                      <p className="text-lg font-black text-slate-900 mt-0.5">{reportData.totalReturnsCount}</p>
                    </div>
                    <div className="h-8 w-px bg-slate-200" />
                    <div className="px-3 py-1 text-center">
                      <p className="text-[10px] uppercase font-extrabold text-slate-500 tracking-wider">PerÃ­odo</p>
                      <p className="text-xs font-bold text-slate-700 mt-1">
                        {new Date(reportRange.start + 'T00:00:00').toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit' })} - {new Date(reportRange.end + 'T00:00:00').toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit' })}
                      </p>
                    </div>
                  </div>
                </div>
              </div>

              {/* Reports Navigation Sub-Tabs */}
              <div className="flex items-center gap-3 bg-white p-2 rounded-2xl border border-slate-200/80 shadow-sm overflow-x-auto">
                <button
                  onClick={() => setReportsTab('overview')}
                  className={`px-5 py-2.5 rounded-xl text-xs sm:text-sm font-black flex items-center gap-2 transition-all shrink-0 ${
                    reportsTab === 'overview'
                      ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white shadow-md shadow-blue-600/20'
                      : 'text-slate-600 hover:text-blue-700 hover:bg-slate-100'
                  }`}
                >
                  <BarChart3 size={17} /> RelatÃ³rios & GrÃ¡ficos
                </button>
                <button
                  onClick={() => setReportsTab('quantitativo')}
                  className={`px-5 py-2.5 rounded-xl text-xs sm:text-sm font-black flex items-center gap-2 transition-all relative shrink-0 ${
                    reportsTab === 'quantitativo'
                      ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white shadow-md shadow-blue-600/20'
                      : 'text-slate-600 hover:text-blue-700 hover:bg-slate-100'
                  }`}
                >
                  <PieChartIcon size={17} /> Quantitativo por Setor
                  <span className="flex items-center gap-1 text-[10px] bg-blue-600 text-white px-2 py-0.5 rounded-full font-bold">
                    Oficial
                  </span>
                </button>
                <button
                  onClick={() => setReportsTab('letterhead')}
                  className={`px-5 py-2.5 rounded-xl text-xs sm:text-sm font-black flex items-center gap-2 transition-all relative shrink-0 ${
                    reportsTab === 'letterhead'
                      ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white shadow-md shadow-blue-600/20'
                      : 'text-slate-600 hover:text-blue-700 hover:bg-slate-100'
                  }`}
                >
                  <FileText size={17} /> Papel Timbrado
                  {letterheadImage ? (
                    <span className="flex items-center gap-1 text-[10px] bg-emerald-500 text-white px-2 py-0.5 rounded-full font-bold">
                      Anexado
                    </span>
                  ) : (
                    <span className="flex items-center gap-1 text-[10px] bg-amber-500 text-white px-2 py-0.5 rounded-full font-bold">
                      Anexar
                    </span>
                  )}
                </button>
              </div>

              {reportsTab === 'overview' && (
                <div className="space-y-8">
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                {/* Print Requests Section - Only for Admin */}
                {isAdmin && (
                  <div className="bg-white p-6 rounded-3xl border border-blue-100/80 shadow-sm hover:shadow-md transition-all duration-300 relative overflow-hidden group">
                    <div className="h-1.5 w-full bg-gradient-to-r from-blue-600 to-indigo-600 absolute top-0 left-0" />
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-5 pt-2">
                      <div className="flex items-center gap-4">
                        <div className="bg-gradient-to-br from-blue-600 to-blue-800 text-white p-3.5 rounded-2xl shadow-md shadow-blue-600/20 group-hover:scale-105 transition-transform">
                          <Printer size={22} />
                        </div>
                        <div>
                          <h3 className="text-base font-black text-slate-900">ImpressÃ£o de SolicitaÃ§Ãµes</h3>
                          <p className="text-slate-500 text-xs font-medium">Imprima as solicitaÃ§Ãµes pendentes e em separaÃ§Ã£o por perÃ­odo</p>
                        </div>
                      </div>
                    </div>

                    <div className="flex flex-col sm:flex-row items-end gap-3 mt-5 pt-4 border-t border-slate-100">
                      <div className="grid grid-cols-2 gap-3 w-full">
                        <div>
                          <label className="block text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1 ml-1">InÃ­cio</label>
                          <input 
                            type="date" 
                            value={printRange.start}
                            onChange={(e) => setPrintRange({...printRange, start: e.target.value})}
                            className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 font-bold text-xs text-slate-800 cursor-pointer"
                          />
                        </div>
                        <div>
                          <label className="block text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1 ml-1">Fim</label>
                          <input 
                            type="date" 
                            value={printRange.end}
                            onChange={(e) => setPrintRange({...printRange, end: e.target.value})}
                            className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 font-bold text-xs text-slate-800 cursor-pointer"
                          />
                        </div>
                      </div>
                      <button 
                        onClick={handlePrintRequests}
                        className="w-full sm:w-auto bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white px-5 py-2.5 rounded-xl font-extrabold text-xs flex items-center justify-center gap-2 hover:from-blue-800 hover:to-indigo-950 transition-all shadow-md shadow-blue-600/20 whitespace-nowrap"
                      >
                        <Printer size={15} /> Imprimir RelatÃ³rio
                      </button>
                    </div>
                  </div>
                )}

                {/* PCA Report Section */}
                {selectedSector === 'Almoxarifado' && (
                  <div className="bg-white p-6 rounded-3xl border border-emerald-100/80 shadow-sm hover:shadow-md transition-all duration-300 relative overflow-hidden group">
                    <div className="h-1.5 w-full bg-gradient-to-r from-emerald-600 to-teal-600 absolute top-0 left-0" />
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-5 pt-2">
                      <div className="flex items-center gap-4">
                        <div className="bg-gradient-to-br from-emerald-600 to-teal-700 text-white p-3.5 rounded-2xl shadow-md shadow-emerald-600/20 group-hover:scale-105 transition-transform">
                          <Calendar size={22} />
                        </div>
                        <div>
                          <h3 className="text-base font-black text-slate-900">RelatÃ³rio PCA</h3>
                          <p className="text-slate-500 text-xs font-medium">Plano Anual de ContrataÃ§Ã£o - Consumo por tipo no perÃ­odo</p>
                        </div>
                      </div>
                    </div>

                    <div className="flex flex-col sm:flex-row items-end gap-3 mt-5 pt-4 border-t border-slate-100">
                      <div className="grid grid-cols-3 gap-2 w-full">
                        <div>
                          <label className="block text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1 ml-1">InÃ­cio</label>
                          <input 
                            type="date" 
                            value={pcaRange.start}
                            onChange={(e) => setPcaRange({...pcaRange, start: e.target.value})}
                            className="w-full px-2.5 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-emerald-500/20 font-bold text-xs text-slate-800 cursor-pointer"
                          />
                        </div>
                        <div>
                          <label className="block text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1 ml-1">Fim</label>
                          <input 
                            type="date" 
                            value={pcaRange.end}
                            onChange={(e) => setPcaRange({...pcaRange, end: e.target.value})}
                            className="w-full px-2.5 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-emerald-500/20 font-bold text-xs text-slate-800 cursor-pointer"
                          />
                        </div>
                        <div>
                          <label className="block text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1 ml-1">Categoria</label>
                          <select 
                            value={pcaCategory}
                            onChange={(e) => setPcaCategory(e.target.value)}
                            className="w-full px-2 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-emerald-500/20 font-bold text-xs text-slate-800 cursor-pointer"
                          >
                            <option value="all">Todas</option>
                            {Object.keys(CATEGORY_COLORS).sort().map(cat => (
                              <option key={cat} value={cat}>{cat}</option>
                            ))}
                          </select>
                        </div>
                      </div>
                      <button 
                        onClick={handleExportPCA}
                        className="w-full sm:w-auto bg-gradient-to-r from-emerald-600 to-teal-700 text-white px-5 py-2.5 rounded-xl font-extrabold text-xs flex items-center justify-center gap-2 hover:from-emerald-700 hover:to-teal-800 transition-all shadow-md shadow-emerald-600/20 whitespace-nowrap"
                      >
                        <Download size={15} /> Gerar PCA
                      </button>
                    </div>
                  </div>
                )}

                {/* Materials Catalog Section - For Leaders */}
                {!isAdmin && (
                  <div className="bg-white p-6 rounded-3xl border border-blue-100/80 shadow-sm hover:shadow-md transition-all duration-300 relative overflow-hidden group lg:col-span-2">
                    <div className="h-1.5 w-full bg-gradient-to-r from-blue-600 to-cyan-500 absolute top-0 left-0" />
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-5 pt-2">
                      <div className="flex items-center gap-4">
                        <div className="bg-gradient-to-br from-blue-600 to-blue-800 text-white p-3.5 rounded-2xl shadow-md shadow-blue-600/20 group-hover:scale-105 transition-transform">
                          <BookOpen size={22} />
                        </div>
                        <div>
                          <h3 className="text-base font-black text-slate-900">DÃºvidas sobre o que pedir?</h3>
                          <p className="text-slate-500 text-xs font-medium">Baixe o catÃ¡logo simplificado contendo todos os nomes dos materiais e categorias cadastradas.</p>
                        </div>
                      </div>
                      <button 
                        onClick={handleExportMaterialsCatalogPDF}
                        className="bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white px-6 py-3 rounded-xl font-extrabold text-xs flex items-center justify-center gap-2 hover:from-blue-800 hover:to-indigo-950 transition-all shadow-md shadow-blue-600/20 whitespace-nowrap"
                      >
                        <Printer size={16} /> Ver CatÃ¡logo de Itens
                      </button>
                    </div>
                  </div>
                )}
              </div>

              {/* KPI Stats Grid */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-5">
                {isAdmin ? (
                  <>
                    {/* Card 1: Entradas */}
                    <div className="bg-white rounded-2xl border border-emerald-100/80 shadow-sm hover:shadow-md hover:border-emerald-200 transition-all duration-300 overflow-hidden group relative">
                      <div className="h-1.5 w-full bg-gradient-to-r from-emerald-500 to-teal-600" />
                      <div className="p-6">
                        <div className="flex items-center justify-between mb-3">
                          <span className="text-xs font-black text-slate-500 uppercase tracking-widest">Entradas no PerÃ­odo</span>
                          <div className="bg-gradient-to-br from-emerald-600 to-teal-700 text-white p-3 rounded-2xl shadow-md shadow-emerald-500/20 group-hover:scale-105 transition-transform">
                            <TrendingUp size={20} />
                          </div>
                        </div>
                        <h3 className="text-3xl font-black text-slate-900 tracking-tight">{reportData.entries.toLocaleString('pt-BR')}</h3>
                        <div className="mt-4 pt-3 border-t border-slate-100 flex items-center justify-between">
                          <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Valor Financeiro</span>
                          <span className="text-xs font-black text-emerald-600">
                            {new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(reportData.entriesValue)}
                          </span>
                        </div>
                      </div>
                    </div>

                    {/* Card 2: SaÃ­das */}
                    <div className="bg-white rounded-2xl border border-rose-100/80 shadow-sm hover:shadow-md hover:border-rose-200 transition-all duration-300 overflow-hidden group relative">
                      <div className="h-1.5 w-full bg-gradient-to-r from-rose-500 to-pink-600" />
                      <div className="p-6">
                        <div className="flex items-center justify-between mb-3">
                          <span className="text-xs font-black text-slate-500 uppercase tracking-widest">SaÃ­das / Consumo</span>
                          <div className="bg-gradient-to-br from-rose-600 to-pink-700 text-white p-3 rounded-2xl shadow-md shadow-rose-500/20 group-hover:scale-105 transition-transform">
                            <TrendingDown size={20} />
                          </div>
                        </div>
                        <h3 className="text-3xl font-black text-slate-900 tracking-tight">{reportData.exits.toLocaleString('pt-BR')}</h3>
                        <div className="mt-4 pt-3 border-t border-slate-100 flex items-center justify-between">
                          <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Valor Baixado</span>
                          <span className="text-xs font-black text-rose-600">
                            {new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(reportData.exitsValue)}
                          </span>
                        </div>
                      </div>
                    </div>

                    {/* Card 3: Valor em Estoque */}
                    <div className="bg-white rounded-2xl border border-indigo-100/80 shadow-sm hover:shadow-md hover:border-indigo-200 transition-all duration-300 overflow-hidden group relative">
                      <div className="h-1.5 w-full bg-gradient-to-r from-indigo-600 to-blue-600" />
                      <div className="p-6">
                        <div className="flex items-center justify-between mb-3">
                          <span className="text-xs font-black text-slate-500 uppercase tracking-widest">PatrimÃ´nio em Saldo</span>
                          <div className="bg-gradient-to-br from-indigo-600 to-blue-700 text-white p-3 rounded-2xl shadow-md shadow-indigo-500/20 group-hover:scale-105 transition-transform">
                            <DollarSign size={20} />
                          </div>
                        </div>
                        <h3 className="text-2xl font-black text-slate-900 tracking-tight">
                          {new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(reportData.totalValue)}
                        </h3>
                        <div className="mt-4 pt-3 border-t border-slate-100 flex items-center gap-1.5">
                          <span className="w-1.5 h-1.5 rounded-full bg-indigo-600" />
                          <span className="text-[10px] font-bold text-indigo-700 uppercase tracking-wider">
                            Valor total ativo
                          </span>
                        </div>
                      </div>
                    </div>

                    {/* Card 4: Itens Ativos */}
                    <div className="bg-white rounded-2xl border border-blue-100/80 shadow-sm hover:shadow-md hover:border-blue-200 transition-all duration-300 overflow-hidden group relative">
                      <div className="h-1.5 w-full bg-gradient-to-r from-blue-600 to-cyan-500" />
                      <div className="p-6">
                        <div className="flex items-center justify-between mb-3">
                          <span className="text-xs font-black text-slate-500 uppercase tracking-widest">Insumos Cadastrados</span>
                          <div className="bg-gradient-to-br from-blue-600 to-cyan-700 text-white p-3 rounded-2xl shadow-md shadow-blue-500/20 group-hover:scale-105 transition-transform">
                            <Package size={20} />
                          </div>
                        </div>
                        <h3 className="text-3xl font-black text-slate-900 tracking-tight">{items.length}</h3>
                        <div className="mt-4 pt-3 border-t border-slate-100 flex items-center gap-1.5">
                          <span className="w-1.5 h-1.5 rounded-full bg-blue-600" />
                          <span className="text-[10px] font-bold text-blue-700 uppercase tracking-wider">
                            Itens no catÃ¡logo
                          </span>
                        </div>
                      </div>
                    </div>
                  </>
                ) : (
                  <>
                    <div className="bg-white p-6 rounded-2xl border border-rose-100 shadow-sm hover:shadow-md transition-all lg:col-span-2 overflow-hidden relative group">
                      <div className="h-1.5 w-full bg-gradient-to-r from-rose-500 to-pink-600 absolute top-0 left-0" />
                      <p className="text-slate-500 text-xs font-black uppercase tracking-wider mb-3">Consumo do Setor no PerÃ­odo</p>
                      <div className="flex items-center gap-5">
                        <div className="bg-rose-50 p-4 rounded-2xl text-rose-600 border border-rose-100 group-hover:scale-105 transition-transform">
                          <ArrowDownLeft size={32} />
                        </div>
                        <div>
                          <h3 className="text-4xl font-black text-rose-600">{reportData.exits.toLocaleString('pt-BR')}</h3>
                          <p className="text-xs font-extrabold text-slate-400 uppercase tracking-widest mt-0.5">Unidades Recebidas</p>
                        </div>
                      </div>
                    </div>
                    <div className="bg-white p-6 rounded-2xl border border-blue-100 shadow-sm hover:shadow-md transition-all lg:col-span-2 overflow-hidden relative group">
                      <div className="h-1.5 w-full bg-gradient-to-r from-blue-600 to-indigo-600 absolute top-0 left-0" />
                      <p className="text-slate-500 text-xs font-black uppercase tracking-wider mb-3">SolicitaÃ§Ãµes no PerÃ­odo</p>
                      <div className="flex items-center gap-5">
                        <div className="bg-blue-50 p-4 rounded-2xl text-blue-600 border border-blue-100 group-hover:scale-105 transition-transform">
                          <FileText size={32} />
                        </div>
                        <div>
                          <h3 className="text-4xl font-black text-blue-600">
                            {requests.filter(r => {
                              const d = new Date(r.date);
                              return r.sector === selectedSector && !r.deletedAt && d >= startOfDay(parseISO(reportRange.start)) && d <= endOfDay(parseISO(reportRange.end));
                            }).length}
                          </h3>
                          <p className="text-xs font-extrabold text-slate-400 uppercase tracking-widest mt-0.5">Pedidos Realizados</p>
                        </div>
                      </div>
                    </div>
                  </>
                )}
              </div>

              {/* Visual Overview Section Header */}
              <div className="flex items-center justify-between px-1 pt-2">
                <div className="flex items-center gap-3">
                  <div className="p-2.5 bg-blue-50 text-blue-700 rounded-2xl border border-blue-100">
                    <BarChart3 size={20} />
                  </div>
                  <div>
                    <h3 className="text-xl font-black text-slate-900 tracking-tight">Panorama Visual de Consumo & MovimentaÃ§Ã£o</h3>
                    <p className="text-xs text-slate-500 font-medium">GrÃ¡ficos interativos para acompanhamento gerencial das operaÃ§Ãµes</p>
                  </div>
                </div>
              </div>

              {/* Charts Grid */}
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                {/* Movement Chart */}
                <div className="bg-white p-6 sm:p-8 rounded-3xl border border-blue-100/80 shadow-sm hover:shadow-md transition-all">
                  <div className="flex items-center justify-between mb-6">
                    <h4 className="text-base font-black text-slate-900 flex items-center gap-2">
                      <Activity size={18} className="text-blue-600" /> MovimentaÃ§Ã£o {isAdmin ? 'Geral' : 'do Setor'}
                    </h4>
                    <span className="text-[10px] font-bold text-slate-400 bg-slate-100 px-2.5 py-1 rounded-lg uppercase tracking-wider">Fluxo DiÃ¡rio</span>
                  </div>
                  <div className="h-[300px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <AreaChart data={reportData.daily}>
                        <defs>
                          <linearGradient id="colorEntries" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="5%" stopColor="#10b981" stopOpacity={0.15}/>
                            <stop offset="95%" stopColor="#10b981" stopOpacity={0}/>
                          </linearGradient>
                          <linearGradient id="colorExits" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="5%" stopColor="#f43f5e" stopOpacity={0.15}/>
                            <stop offset="95%" stopColor="#f43f5e" stopOpacity={0}/>
                          </linearGradient>
                        </defs>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                        <XAxis dataKey="date" axisLine={false} tickLine={false} tick={{fontSize: 10, fill: '#64748b', fontWeight: 600}} />
                        <YAxis axisLine={false} tickLine={false} tick={{fontSize: 10, fill: '#64748b', fontWeight: 600}} />
                        <Tooltip 
                          contentStyle={{ borderRadius: '16px', border: '1px solid #e2e8f0', boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.08)', fontWeight: 'bold' }}
                        />
                        {isAdmin && <Area type="monotone" dataKey="entries" name="Entradas" stroke="#10b981" fillOpacity={1} fill="url(#colorEntries)" strokeWidth={3} />}
                        <Area type="monotone" dataKey="exits" name={isAdmin ? "SaÃ­das" : "Consumo"} stroke="#f43f5e" fillOpacity={1} fill="url(#colorExits)" strokeWidth={3} />
                      </AreaChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                {/* Category Breakdown */}
                <div className="bg-white p-6 sm:p-8 rounded-3xl border border-blue-100/80 shadow-sm hover:shadow-md transition-all">
                  <div className="flex items-center justify-between mb-6">
                    <h4 className="text-base font-black text-slate-900 flex items-center gap-2">
                      <PieChartIcon size={18} className="text-amber-500" /> DistribuiÃ§Ã£o de Consumo por Categoria
                    </h4>
                    <span className="text-[10px] font-bold text-slate-400 bg-slate-100 px-2.5 py-1 rounded-lg uppercase tracking-wider">ProporÃ§Ã£o</span>
                  </div>
                  <div className="h-[300px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <PieChart>
                        <Pie
                          data={reportData.consumptionCategories}
                          cx="50%"
                          cy="50%"
                          innerRadius={60}
                          outerRadius={95}
                          paddingAngle={4}
                          dataKey="value"
                          label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`}
                        >
                          {reportData.consumptionCategories.map((entry, index) => (
                            <Cell key={`cell-${index}`} fill={getCategoryColor(entry.name)} />
                          ))}
                        </Pie>
                        <Tooltip 
                          contentStyle={{ borderRadius: '16px', border: '1px solid #e2e8f0', boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.08)', fontWeight: 'bold' }}
                        />
                        <Legend verticalAlign="bottom" height={36}/>
                      </PieChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                {/* Top Consumed Items - Only for Admin */}
                {isAdmin && (
                  <div className="bg-white p-6 sm:p-8 rounded-3xl border border-blue-100/80 shadow-sm hover:shadow-md transition-all lg:col-span-2">
                    <div className="flex items-center justify-between mb-6">
                      <h4 className="text-base font-black text-slate-900 flex items-center gap-2">
                        <ArrowDownLeft size={18} className="text-rose-600" /> Ranking: Itens Mais Consumidos
                      </h4>
                      <span className="text-[10px] font-bold text-rose-600 bg-rose-50 px-2.5 py-1 rounded-lg uppercase tracking-wider">Top Demandas</span>
                    </div>
                    <div className="h-[320px] w-full">
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={reportData.topConsumed} layout="vertical">
                          <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                          <XAxis type="number" axisLine={false} tickLine={false} tick={{fontSize: 10, fill: '#64748b'}} />
                          <YAxis dataKey="name" type="category" axisLine={false} tickLine={false} tick={{fontSize: 10, fill: '#1e293b', fontWeight: 'bold'}} width={130} />
                          <Tooltip 
                            cursor={{fill: '#f8fafc'}}
                            contentStyle={{ borderRadius: '16px', border: '1px solid #e2e8f0', boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.08)', fontWeight: 'bold' }}
                          />
                          <Bar dataKey="value" name="Qtd Consumida" fill="#f43f5e" radius={[0, 8, 8, 0]} barSize={18} />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                )}

                {/* Only for Admin Charts */}
                {isAdmin && (
                  <>
                    {/* Stock Value by Category */}
                    <div className="bg-white p-6 sm:p-8 rounded-3xl border border-blue-100/80 shadow-sm hover:shadow-md transition-all">
                      <div className="flex items-center justify-between mb-6">
                        <h4 className="text-base font-black text-slate-900 flex items-center gap-2">
                          <DollarSign size={18} className="text-emerald-600" /> Valor em Estoque por Categoria
                        </h4>
                        <span className="text-[10px] font-bold text-emerald-600 bg-emerald-50 px-2.5 py-1 rounded-lg uppercase tracking-wider">Financeiro</span>
                      </div>
                      <div className="h-[300px] w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={reportData.categoryValues} layout="vertical">
                            <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                            <XAxis type="number" axisLine={false} tickLine={false} tick={{fontSize: 10, fill: '#64748b'}} />
                            <YAxis dataKey="name" type="category" axisLine={false} tickLine={false} tick={{fontSize: 10, fill: '#1e293b', fontWeight: 'bold'}} width={130} />
                            <Tooltip 
                              cursor={{fill: '#f8fafc'}}
                              formatter={(value: number) => new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value)}
                              contentStyle={{ borderRadius: '16px', border: '1px solid #e2e8f0', boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.08)', fontWeight: 'bold' }}
                            />
                            <Bar dataKey="value" name="Valor Total">
                              {reportData.categoryValues.map((entry, index) => (
                                <Cell key={`cell-${index}`} fill={getCategoryColor(entry.name)} />
                              ))}
                            </Bar>
                            <Legend />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    </div>

                    {/* Exits by Reason */}
                    <div className="bg-white p-6 sm:p-8 rounded-3xl border border-blue-100/80 shadow-sm hover:shadow-md transition-all">
                      <div className="flex items-center justify-between mb-6">
                        <h4 className="text-base font-black text-slate-900 flex items-center gap-2">
                          <TrendingDown size={18} className="text-rose-600" /> SaÃ­das por Motivo
                        </h4>
                        <span className="text-[10px] font-bold text-slate-400 bg-slate-100 px-2.5 py-1 rounded-lg uppercase tracking-wider">DestinaÃ§Ã£o</span>
                      </div>
                      <div className="h-[300px] w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <PieChart>
                            <Pie
                              data={[
                                { name: 'Consumo', value: reportData.exitsByReason.consumo },
                                { name: 'DoaÃ§Ã£o', value: reportData.exitsByReason.doacao },
                                { name: 'Vencimento', value: reportData.exitsByReason.vencido },
                                { name: 'Perda/Avaria', value: reportData.exitsByReason.perda || 0 }
                              ]}
                              cx="50%"
                              cy="50%"
                              innerRadius={60}
                              outerRadius={95}
                              paddingAngle={4}
                              dataKey="value"
                              label={({ name, value }) => `${name}: ${value}`}
                            >
                              <Cell fill="#2563eb" />
                              <Cell fill="#f59e0b" />
                              <Cell fill="#ef4444" />
                              <Cell fill="#64748b" />
                            </Pie>
                            <Tooltip 
                              contentStyle={{ borderRadius: '16px', border: '1px solid #e2e8f0', boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.08)', fontWeight: 'bold' }}
                            />
                            <Legend verticalAlign="bottom" height={36}/>
                          </PieChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                  </>
                )}

                {/* Exits by Sector - Only for Admin */}
                {isAdmin && (
                  <div className="bg-white p-6 sm:p-8 rounded-3xl border border-blue-100/80 shadow-sm hover:shadow-md transition-all">
                    <div className="flex items-center justify-between mb-6">
                      <h4 className="text-base font-black text-slate-900 flex items-center gap-2">
                        <ArrowUpRight size={18} className="text-rose-600" /> SaÃ­das por Setor (Quantidade por Tipo)
                      </h4>
                      <span className="text-[10px] font-bold text-slate-400 bg-slate-100 px-2.5 py-1 rounded-lg uppercase tracking-wider">Setorial</span>
                    </div>
                    <div className="h-[300px] w-full">
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={reportData.sectors} layout="vertical">
                          <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                          <XAxis type="number" axisLine={false} tickLine={false} tick={{fontSize: 10, fill: '#64748b'}} />
                          <YAxis dataKey="name" type="category" axisLine={false} tickLine={false} tick={{fontSize: 10, fill: '#1e293b', fontWeight: 'bold'}} width={110} />
                          <Tooltip 
                            cursor={{fill: '#f8fafc'}}
                            contentStyle={{ borderRadius: '16px', border: '1px solid #e2e8f0', boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.08)', fontWeight: 'bold' }}
                          />
                          {reportData.categoriesInSector.map((cat: string) => (
                            <Bar 
                              key={cat} 
                              dataKey={cat} 
                              name={cat} 
                              stackId="a" 
                              fill={getCategoryColor(cat)} 
                              radius={[0, 0, 0, 0]} 
                              barSize={18} 
                            />
                          ))}
                          <Legend />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                )}

                {/* Returns by Sector - DevoluÃ§Ãµes por Setor */}
                <div className="bg-white p-6 sm:p-8 rounded-3xl border border-amber-100/80 shadow-sm hover:shadow-md transition-all">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-6">
                    <div>
                      <h4 className="text-base font-black text-slate-900 flex items-center gap-2">
                        <RotateCcw size={18} className="text-amber-600" /> DevoluÃ§Ãµes por Setor
                      </h4>
                      <p className="text-xs text-slate-500 font-medium">Materiais devolvidos ao almoxarifado pelos setores</p>
                    </div>
                    <div className="flex items-center gap-2 shrink-0">
                      <span className="text-[10px] font-black text-amber-700 bg-amber-50 border border-amber-200/80 px-2.5 py-1 rounded-lg uppercase tracking-wider">
                        {reportData.totalReturnsCount} {reportData.totalReturnsCount === 1 ? 'Item Devolvido' : 'Itens Devolvidos'}
                      </span>
                      {isAdmin && reportData.totalReturnsValue > 0 && (
                        <span className="text-[10px] font-black text-emerald-700 bg-emerald-50 border border-emerald-200/80 px-2.5 py-1 rounded-lg uppercase tracking-wider">
                          {new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(reportData.totalReturnsValue)}
                        </span>
                      )}
                    </div>
                  </div>

                  {reportData.returnsBySector.length > 0 ? (
                    <div className="h-[300px] w-full">
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={reportData.returnsBySector} layout="vertical">
                          <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                          <XAxis type="number" axisLine={false} tickLine={false} tick={{fontSize: 10, fill: '#64748b'}} />
                          <YAxis dataKey="name" type="category" axisLine={false} tickLine={false} tick={{fontSize: 10, fill: '#1e293b', fontWeight: 'bold'}} width={110} />
                          <Tooltip 
                            cursor={{fill: '#fffbeb'}}
                            formatter={(value: number, name: string) => [
                              name === 'Qtd Devolvida' ? `${value} un.` : new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value),
                              'Quantidade'
                            ]}
                            contentStyle={{ borderRadius: '16px', border: '1px solid #fde68a', boxShadow: '0 10px 25px -5px rgba(245, 158, 11, 0.1)', fontWeight: 'bold' }}
                          />
                          <Bar dataKey="quantity" name="Qtd Devolvida" fill="#d97706" radius={[0, 8, 8, 0]} barSize={18} />
                          <Legend />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  ) : (
                    <div className="h-[220px] flex flex-col items-center justify-center text-center p-6 bg-amber-50/30 rounded-2xl border border-dashed border-amber-200/60">
                      <div className="w-10 h-10 rounded-2xl bg-amber-100/80 text-amber-700 flex items-center justify-center mb-2 shadow-xs">
                        <RotateCcw size={20} />
                      </div>
                      <p className="text-sm font-black text-slate-800">Nenhuma devoluÃ§Ã£o registrada no perÃ­odo</p>
                      <p className="text-xs text-slate-500 max-w-sm mt-1 font-medium">Os materiais que forem devolvidos pelos setores ao almoxarifado no perÃ­odo selecionado aparecerÃ£o consolidados neste grÃ¡fico.</p>
                    </div>
                  )}
                </div>

                {/* Value by Supplier */}
                {isAdmin && (
                  <div className="bg-white p-6 sm:p-8 rounded-3xl border border-blue-100/80 shadow-sm hover:shadow-md transition-all">
                    <div className="flex items-center justify-between mb-6">
                      <h4 className="text-base font-black text-slate-900 flex items-center gap-2">
                        <DollarSign size={18} className="text-amber-500" /> Valor por Fornecedor
                      </h4>
                      <span className="text-[10px] font-bold text-amber-600 bg-amber-50 px-2.5 py-1 rounded-lg uppercase tracking-wider">Fornecedores</span>
                    </div>
                    <div className="h-[300px] w-full">
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={reportData.suppliers} layout="vertical" margin={{ left: 10, right: 20 }}>
                          <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                          <XAxis type="number" axisLine={false} tickLine={false} tick={{fontSize: 10, fill: '#64748b'}} />
                          <YAxis 
                            dataKey="name" 
                            type="category" 
                            axisLine={false} 
                            tickLine={false} 
                            tick={{fontSize: 10, fill: '#1e293b', fontWeight: 'bold'}} 
                            width={140}
                          />
                          <Tooltip 
                            cursor={{fill: '#f8fafc'}}
                            formatter={(value: number) => new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value)}
                            contentStyle={{ borderRadius: '16px', border: '1px solid #e2e8f0', boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.08)', fontWeight: 'bold' }}
                          />
                          <Bar dataKey="value" name="Valor Total" fill="#f59e0b" radius={[0, 8, 8, 0]} barSize={18} />
                          <Legend />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                )}

                {/* Top Requested Items - Only for Admin */}
                {isAdmin && (
                  <div className="bg-white p-6 sm:p-8 rounded-3xl border border-blue-100/80 shadow-sm hover:shadow-md transition-all lg:col-span-2">
                    <div className="flex items-center justify-between mb-6">
                      <h4 className="text-base font-black text-slate-900 flex items-center gap-2">
                        <Plus size={18} className="text-blue-600" /> Itens Mais Solicitados (Top 10)
                      </h4>
                      <span className="text-[10px] font-bold text-blue-600 bg-blue-50 px-2.5 py-1 rounded-lg uppercase tracking-wider">Pedidos</span>
                    </div>
                    <div className="h-[320px] w-full">
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={reportData.topRequested}>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                          <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{fontSize: 10, fill: '#1e293b', fontWeight: 'bold'}} />
                          <YAxis axisLine={false} tickLine={false} tick={{fontSize: 10, fill: '#64748b'}} />
                          <Tooltip 
                            cursor={{fill: '#f8fafc'}}
                            contentStyle={{ borderRadius: '16px', border: '1px solid #e2e8f0', boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.08)', fontWeight: 'bold' }}
                          />
                          <Bar dataKey="value" name="Qtd Solicitada" fill="#2563eb" radius={[8, 8, 0, 0]} barSize={36} />
                          <Legend />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                )}

                {/* Extra vs Contract Comparison */}
                {isAdmin && (
                  <div className="bg-white p-6 sm:p-8 rounded-3xl border border-blue-100/80 shadow-sm hover:shadow-md transition-all lg:col-span-2">
                    <div className="flex items-center justify-between mb-6">
                      <h4 className="text-base font-black text-slate-900 flex items-center gap-2">
                        <BarChart3 size={18} className="text-indigo-600" /> Comparativo: Contrato vs Extra vs DoaÃ§Ã£o
                      </h4>
                      <span className="text-[10px] font-bold text-indigo-600 bg-indigo-50 px-2.5 py-1 rounded-lg uppercase tracking-wider">Origem</span>
                    </div>
                    <div className="h-[300px] w-full">
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart 
                          data={[
                            { 
                              name: 'Entradas', 
                              contrato: reportData.originStats.contract.entries, 
                              extra: reportData.originStats.extra.entries,
                              doacao: reportData.originStats.donation.entries
                            },
                            { 
                              name: 'SaÃ­das', 
                              contrato: reportData.originStats.contract.exits, 
                              extra: reportData.originStats.extra.exits,
                              doacao: reportData.originStats.donation.exits
                            },
                            { 
                              name: 'Estoque Atual', 
                              contrato: reportData.originStats.contract.current, 
                              extra: reportData.originStats.extra.current,
                              doacao: reportData.originStats.donation.current
                            }
                          ]}
                        >
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                          <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{fontSize: 11, fill: '#1e293b', fontWeight: 'bold'}} />
                          <YAxis axisLine={false} tickLine={false} tick={{fontSize: 10, fill: '#64748b'}} />
                          <Tooltip 
                            cursor={{fill: '#f8fafc'}}
                            contentStyle={{ borderRadius: '16px', border: '1px solid #e2e8f0', boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.08)', fontWeight: 'bold' }}
                          />
                          <Legend />
                          <Bar dataKey="contrato" name="Contrato" fill="#1e293b" radius={[6, 6, 0, 0]} barSize={26} />
                          <Bar dataKey="extra" name="Extra" fill="#6366f1" radius={[6, 6, 0, 0]} barSize={26} />
                          <Bar dataKey="doacao" name="DoaÃ§Ã£o" fill="#10b981" radius={[6, 6, 0, 0]} barSize={26} />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                )}
              </div>

              {/* Detailed Sector Breakdown - Visible for Admin and Sector Leaders */}
              {(isAdmin || userProfile?.role === 'SETOR' || userProfile?.role === 'LÃDER') && (
                <div className="bg-white p-6 sm:p-8 rounded-3xl border border-blue-100/80 shadow-sm lg:col-span-2">
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-8 pb-6 border-b border-slate-100">
                    <div>
                      <h4 className="text-lg font-black text-slate-900 flex items-center gap-2 mb-1">
                        <History size={20} className="text-blue-600" /> 
                        RelatÃ³rio Detalhado de Consumo por Item
                      </h4>
                      <p className="text-xs text-slate-500 font-medium">
                        {isAdmin ? (reportSectorFilter === 'all' ? 'Todos os Setores' : `Setor: ${reportSectorFilter}`) : `Setor: ${selectedSector}`} - {format(parseISO(reportRange.start), 'dd/MM/yyyy')} a {format(parseISO(reportRange.end), 'dd/MM/yyyy')}
                      </p>
                    </div>
                    <div className="flex flex-wrap items-center gap-3">
                      {isAdmin && (
                        <div className="text-right mr-2 hidden sm:block">
                          <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Total de SaÃ­das</p>
                          <p className="text-xl font-black text-rose-600">
                            {new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(reportData.consumptionReport.reduce((sum, i) => sum + i.totalValue, 0))}
                          </p>
                        </div>
                      )}
                      <div className="flex flex-wrap gap-2">
                        {isAdmin && (
                          <button 
                            onClick={() => {
                              setSelectedRoomCategories([...categories]);
                              setShowRoomInventoryModal(true);
                            }}
                            className="flex items-center gap-2 bg-blue-600 text-white px-4 py-2.5 rounded-xl text-xs font-bold hover:bg-blue-700 transition-all shadow-md shadow-blue-500/10 active:scale-95"
                          >
                            <Printer size={15} /> Mapa de Sala (Porta)
                          </button>
                        )}
                        <button 
                          onClick={handleExportConsumptionPDF}
                          className="flex items-center gap-2 bg-gradient-to-r from-blue-700 to-indigo-900 text-white px-4 py-2.5 rounded-xl text-xs font-bold hover:from-blue-800 hover:to-indigo-950 transition-all shadow-md shadow-blue-900/10 active:scale-95"
                        >
                          <Download size={15} /> Exportar PDF Consumo
                        </button>
                      </div>
                    </div>
                  </div>
                  
                  <div className="overflow-x-auto rounded-2xl border border-slate-100">
                    <table className="w-full text-left border-collapse min-w-[650px]">
                      <thead>
                        <tr className="bg-slate-50/80 border-b border-slate-200/80">
                          <th className="py-3.5 px-5 font-black text-xs text-slate-500 uppercase tracking-wider">Setor / Item</th>
                          <th className="py-3.5 px-4 font-black text-xs text-slate-500 uppercase tracking-wider">Categoria</th>
                          <th className="py-3.5 px-4 font-black text-xs text-slate-500 uppercase tracking-wider text-center">Quantidade</th>
                          {isAdmin && <th className="py-3.5 px-5 font-black text-xs text-slate-500 uppercase tracking-wider text-right">Valor Total</th>}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {reportData.consumptionBySector.map((sectorGroup, idx) => (
                          <React.Fragment key={idx}>
                            <tr className="bg-blue-50/40 border-y border-blue-100/60">
                              <td className="py-2.5 px-5 font-black text-[11px] uppercase tracking-wider text-blue-900 flex items-center gap-2" colSpan={isAdmin ? 3 : 3}>
                                <span className="w-2 h-2 rounded-full bg-blue-600"></span>
                                {sectorGroup.sector}
                              </td>
                              {isAdmin && (
                                <td className="py-2.5 px-5 text-right font-black text-blue-950 text-xs">
                                  {new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(sectorGroup.totalValue)}
                                </td>
                              )}
                            </tr>
                            {Object.values(sectorGroup.items).sort((a, b) => b.quantity - a.quantity).map((item, iIdx) => (
                              <tr key={`${idx}-${iIdx}`} className="hover:bg-blue-50/20 transition-all border-b border-slate-100/80 last:border-b-0">
                                <td className="py-3.5 px-8 text-sm font-semibold text-slate-800">
                                  {item.name}
                                </td>
                                <td className="py-3.5 px-4">
                                  <span 
                                    className="text-[10px] font-black px-2.5 py-1 rounded-md text-white whitespace-nowrap shadow-xs"
                                    style={{ backgroundColor: getCategoryColor(item.category) }}
                                  >
                                    {item.category}
                                  </span>
                                </td>
                                <td className="py-3.5 px-4 text-center">
                                  <span className="text-slate-900 font-extrabold text-sm bg-slate-100 px-3 py-1 rounded-lg border border-slate-200/60">
                                    {item.quantity}
                                  </span>
                                </td>
                                {isAdmin && (
                                  <td className="py-3.5 px-5 text-right font-bold text-slate-600 text-sm">
                                    {new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(item.value)}
                                  </td>
                                )}
                              </tr>
                            ))}
                          </React.Fragment>
                        ))}
                        {reportData.consumptionBySector.length === 0 && (
                          <tr>
                            <td colSpan={isAdmin ? 4 : 3} className="py-12 text-center text-slate-400 font-medium italic">
                              Nenhuma saÃ­da registrada para este perÃ­odo ou setor selecionado.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
              </div>
              )}

              {reportsTab === 'quantitativo' && (
                <div className="space-y-6">
                  {/* Action & Filter Controls Bar */}
                  <div className="bg-white p-5 sm:p-6 rounded-3xl border border-slate-200/90 shadow-xs space-y-4">
                    <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 pb-4 border-b border-slate-100">
                      <div>
                        <div className="flex items-center gap-2">
                          <span className="px-3 py-1 rounded-full text-xs font-extrabold bg-blue-50 text-blue-700 border border-blue-200">
                            RelatÃ³rio Oficial DispensaÃ§Ã£o
                          </span>
                          <span className="px-3 py-1 rounded-full text-xs font-extrabold bg-emerald-50 text-emerald-700 border border-emerald-200">
                            {quantitativoSource === 'sample' ? 'Exemplo Oficial Sobral' : 'Dados do Sistema'}
                          </span>
                        </div>
                        <h3 className="text-xl sm:text-2xl font-black text-slate-900 mt-2">
                          Quantitativo de Materiais por Setor
                        </h3>
                        <p className="text-xs sm:text-sm text-slate-500 font-medium mt-1">
                          Gere o documento oficial com grÃ¡fico e anÃ¡lise crÃ­tica para apresentaÃ§Ã£o gerencial e fiscal referente Ã  categoria selecionada.
                        </p>
                      </div>

                      {/* Export Buttons */}
                      <div className="flex flex-wrap items-center gap-2.5 shrink-0">
                        <button
                          onClick={() => handleExportQuantitativoExcel()}
                          className="px-4 py-2.5 rounded-xl bg-slate-100 text-slate-800 font-extrabold text-xs flex items-center gap-2 hover:bg-slate-200 transition-all border border-slate-200/80"
                        >
                          <Download size={15} /> Excel (.xlsx)
                        </button>
                        <button
                          onClick={() => setIsEditingQuantitativoAnalysis(!isEditingQuantitativoAnalysis)}
                          className="px-4 py-2.5 rounded-xl bg-slate-900 text-white font-extrabold text-xs flex items-center gap-2 hover:bg-slate-800 transition-all shadow-xs"
                        >
                          <Edit2 size={15} /> {isEditingQuantitativoAnalysis ? 'Concluir EdiÃ§Ã£o' : 'Editar AnÃ¡lise CrÃ­tica'}
                        </button>
                        <button
                          onClick={handleExportQuantitativoPDF}
                          className="px-5 py-2.5 rounded-xl bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white font-black text-xs flex items-center gap-2 hover:from-blue-800 hover:to-indigo-950 transition-all shadow-md shadow-blue-600/20"
                        >
                          <Printer size={15} /> Exportar PDF Oficial
                        </button>
                      </div>
                    </div>

                    {/* Filter Parameters */}
                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 pt-1">
                      <div>
                        <label className="block text-[10px] font-black text-slate-500 uppercase tracking-wider mb-1">
                          Origem dos Dados
                        </label>
                        <select
                          value={quantitativoSource}
                          onChange={(e) => setQuantitativoSource(e.target.value as 'sample' | 'system')}
                          className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-bold text-xs text-slate-800 focus:ring-2 focus:ring-blue-500/20"
                        >
                          <option value="system">Dados Reais do Sistema (PadrÃ£o)</option>
                          <option value="sample">Exemplo Demonstrativo</option>
                        </select>
                      </div>

                      <div>
                        <label className="block text-[10px] font-black text-slate-500 uppercase tracking-wider mb-1">
                          PerÃ­odo de ReferÃªncia
                        </label>
                        <select
                          value={quantitativoPeriodPreset}
                          onChange={(e) => setQuantitativoPeriodPreset(e.target.value as any)}
                          className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-bold text-xs text-slate-800 focus:ring-2 focus:ring-blue-500/20"
                        >
                          <option value="1_semestre_2026">1Âº Semestre de 2026 (Jan - Jun)</option>
                          <option value="2_semestre_2026">2Âº Semestre de 2026 (Jul - Dez)</option>
                          <option value="ano_2026">Ano Completo de 2026 (Total)</option>
                          <option value="custom">PerÃ­odo Personalizado</option>
                        </select>
                      </div>

                      <div>
                        <label className="block text-[10px] font-black text-slate-500 uppercase tracking-wider mb-1">
                          Categoria de Materiais
                        </label>
                        <select
                          value={quantitativoCategory}
                          onChange={(e) => {
                            setQuantitativoCategory(e.target.value);
                            setQuantitativoTitle('');
                          }}
                          className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-bold text-xs text-slate-800 focus:ring-2 focus:ring-blue-500/20"
                        >
                          <option value="Material MÃ©dico-Hospitalar">Material MÃ©dico-Hospitalar</option>
                          <option value="Medicamentos">Medicamentos</option>
                          <option value="AlimentÃ­cio">AlimentÃ­cio</option>
                          <option value="Expediente">Expediente / Papelaria</option>
                          <option value="Higiene e Limpeza">Higiene e Limpeza</option>
                          <option value="OdontolÃ³gico">OdontolÃ³gico</option>
                          <option value="RadiolÃ³gico">RadiolÃ³gico</option>
                          <option value="EPI e SeguranÃ§a">EPI e SeguranÃ§a</option>
                          <option value="InformÃ¡tica">InformÃ¡tica / TI</option>
                          <option value="Copa & Cozinha">Copa & Cozinha</option>
                          <option value="ManutenÃ§Ã£o">ManutenÃ§Ã£o</option>
                          <option value="Todos">Todos os Materiais (Total Geral)</option>
                        </select>
                      </div>

                      <div>
                        <label className="block text-[10px] font-black text-slate-500 uppercase tracking-wider mb-1">
                          TÃ­tulo do Documento
                        </label>
                        <input
                          type="text"
                          value={quantitativoTitle}
                          onChange={(e) => setQuantitativoTitle(e.target.value)}
                          className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-bold text-xs text-slate-800 focus:ring-2 focus:ring-blue-500/20"
                        />
                      </div>

                      <div className="sm:col-span-2 lg:col-span-4">
                        <div className="flex items-center justify-between mb-1">
                          <label className="block text-[10px] font-black text-slate-500 uppercase tracking-wider">
                            Texto da AnÃ¡lise CrÃ­tica (Gerada pelo GrÃ¡fico / EditÃ¡vel)
                          </label>
                          <button
                            type="button"
                            onClick={() => setQuantitativoCriticalAnalysis('')}
                            className="text-[10px] font-bold text-blue-700 hover:underline cursor-pointer flex items-center gap-1"
                          >
                            <RotateCcw size={10} /> Recalcular AutomÃ¡tico pelo GrÃ¡fico
                          </button>
                        </div>
                        <textarea
                          rows={3}
                          value={quantitativoCriticalAnalysis !== '' ? quantitativoCriticalAnalysis : quantitativoReportData.criticalAnalysis}
                          onChange={(e) => setQuantitativoCriticalAnalysis(e.target.value)}
                          placeholder="Digite ou edite o texto da AnÃ¡lise CrÃ­tica do relatÃ³rio..."
                          className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-medium text-xs text-slate-800 focus:ring-2 focus:ring-blue-500/20 leading-relaxed"
                        />
                      </div>
                    </div>
                  </div>

                  {/* Printable Document A4 Canvas Container (Landscape Orientation) */}
                  <div className="bg-slate-200/80 p-4 sm:p-8 rounded-3xl border border-slate-300 flex justify-center shadow-inner overflow-x-auto">
                    <div
                      ref={quantitativoReportRef}
                      className="bg-white w-full max-w-[1120px] p-8 sm:p-12 shadow-2xl rounded-xl border border-slate-300 text-slate-900 space-y-6 relative font-sans shrink-0"
                      style={{ minWidth: '920px' }}
                    >
                      {/* Document Header - Timbrado Image Only */}
                      <div className="pb-4 border-b-2 border-slate-200 flex justify-center items-center min-h-[70px]">
                        <img 
                          src={letterheadImage || "/official_letterhead.svg"} 
                          alt="Papel Timbrado Oficial" 
                          className="w-full max-h-24 object-contain" 
                          onError={(e) => {
                            const logoToUse = appRectangularLogo || appLogo;
                            if (logoToUse) {
                              (e.target as HTMLElement).setAttribute('src', logoToUse);
                            } else {
                              (e.target as HTMLElement).style.display = 'none';
                            }
                          }}
                        />
                      </div>

                      {/* Document Title */}
                      <div className="text-center py-2">
                        <h1 className="text-sm sm:text-base font-black text-slate-950 uppercase tracking-tight leading-snug max-w-4xl mx-auto">
                          {quantitativoReportData.title}
                        </h1>
                      </div>

                      {/* Stacked Bar Chart Matrix */}
                      <div className="space-y-2 py-2">
                        {/* Row Headers & Bars */}
                        {quantitativoReportData.sectors.map((sec, idx) => (
                          <div key={idx} className="flex items-center gap-3">
                            <div className="w-36 sm:w-44 text-right shrink-0">
                              <span className="text-[11px] font-extrabold text-slate-800 uppercase tracking-tight truncate block">
                                {sec.name}
                              </span>
                            </div>

                            {/* Stacked Bar Track */}
                            <div className="flex-1 h-6 bg-slate-100 border border-slate-300 rounded-sm overflow-hidden flex relative shadow-2xs">
                              {sec.values.map((val, mIdx) => {
                                if (val === 0 || sec.total === 0) return null;
                                const pct = (val / sec.total) * 100;
                                return (
                                  <div
                                    key={mIdx}
                                    className="h-full flex items-center justify-center text-[10px] font-black text-white px-1 overflow-hidden transition-all"
                                    style={{
                                      width: `${pct}%`,
                                      backgroundColor: quantitativoReportData.monthColors[mIdx % quantitativoReportData.monthColors.length]
                                    }}
                                    title={`${quantitativoReportData.months[mIdx]}: ${val}`}
                                  >
                                    {pct >= 4 ? val : ''}
                                  </div>
                                );
                              })}
                            </div>

                            {/* Total Geral Badge */}
                            <div className="w-16 text-right shrink-0">
                              <span className="px-2 py-0.5 rounded bg-slate-900 text-white font-extrabold text-xs shadow-2xs inline-block text-center w-full">
                                {sec.total}
                              </span>
                            </div>
                          </div>
                        ))}

                        {/* Scale Axis % */}
                        <div className="flex items-center justify-between text-[9px] font-bold text-slate-400 pt-2 px-36 sm:px-44">
                          <span>0%</span>
                          <span>25%</span>
                          <span>50%</span>
                          <span>75%</span>
                          <span>100%</span>
                        </div>

                        {/* Month Legend Bar */}
                        <div className="flex flex-wrap items-center justify-center gap-3 pt-3 border-t border-slate-200">
                          {quantitativoReportData.months.map((m, mIdx) => (
                            <div key={mIdx} className="flex items-center gap-1.5">
                              <span
                                className="w-3.5 h-3.5 rounded-xs inline-block shadow-2xs"
                                style={{ backgroundColor: quantitativoReportData.monthColors[mIdx % quantitativoReportData.monthColors.length] }}
                              />
                              <span className="text-[11px] font-extrabold text-slate-700">{m}</span>
                            </div>
                          ))}
                          <div className="flex items-center gap-1.5 ml-2">
                            <span className="w-3.5 h-3.5 rounded-xs bg-slate-900 inline-block shadow-2xs" />
                            <span className="text-[11px] font-extrabold text-slate-900">Total geral</span>
                          </div>
                        </div>
                      </div>

                      {/* AnÃ¡lise CrÃ­tica Section */}
                      <div className="pt-3 border-t-2 border-slate-200 space-y-2">
                        <div className="flex items-center justify-between">
                          <h3 className="text-xs font-black uppercase tracking-wider flex items-center gap-2" style={{ color: '#0f172a' }}>
                            <BarChart3 size={15} style={{ color: '#334155' }} />
                            AnÃ¡lise CrÃ­tica:
                          </h3>
                          <button
                            data-pdf-hide="true"
                            type="button"
                            onClick={() => setIsEditingQuantitativoAnalysis(!isEditingQuantitativoAnalysis)}
                            className="text-[10px] font-bold text-slate-700 hover:text-slate-900 flex items-center gap-1 cursor-pointer bg-slate-100 hover:bg-slate-200 px-2.5 py-1 rounded-lg border border-slate-300 transition-all print:hidden"
                          >
                            <Edit2 size={12} />
                            {isEditingQuantitativoAnalysis ? 'Salvar EdiÃ§Ã£o' : 'Editar AnÃ¡lise'}
                          </button>
                        </div>

                        {isEditingQuantitativoAnalysis ? (
                          <div className="space-y-2">
                            <textarea
                              rows={6}
                              value={quantitativoCriticalAnalysis !== '' ? quantitativoCriticalAnalysis : quantitativoReportData.criticalAnalysis}
                              onChange={(e) => setQuantitativoCriticalAnalysis(e.target.value)}
                              placeholder="Digite ou edite o texto da AnÃ¡lise CrÃ­tica..."
                              className="w-full p-3 border border-slate-300 rounded-xl text-xs font-medium leading-relaxed focus:ring-2 focus:ring-slate-400/20"
                              style={{ backgroundColor: '#f8fafc', color: '#0f172a', borderColor: '#cbd5e1' }}
                            />
                            <div data-pdf-hide="true" className="text-[10px] text-slate-500 font-bold flex flex-wrap justify-between items-center gap-2 print:hidden">
                              <span>* O texto acima serÃ¡ impresso no relatÃ³rio oficial em PDF.</span>
                              <div className="flex items-center gap-3">
                                <button
                                  type="button"
                                  onClick={() => setQuantitativoCriticalAnalysis('')}
                                  className="text-slate-700 hover:underline flex items-center gap-1 font-bold"
                                >
                                  <RotateCcw size={10} /> Recalcular pelo GrÃ¡fico
                                </button>
                                <button
                                  type="button"
                                  onClick={() => setIsEditingQuantitativoAnalysis(false)}
                                  className="text-slate-900 underline font-black hover:text-slate-950"
                                >
                                  Concluir EdiÃ§Ã£o
                                </button>
                              </div>
                            </div>
                          </div>
                        ) : (
                          <div
                            onClick={() => setIsEditingQuantitativoAnalysis(true)}
                            title="Clique para editar o texto da AnÃ¡lise CrÃ­tica"
                            className="group cursor-pointer relative"
                          >
                            <p 
                              className="text-xs font-medium leading-relaxed text-justify p-4 rounded-xl border border-slate-200 group-hover:border-slate-400 transition-colors"
                              style={{ backgroundColor: '#f8fafc', color: '#0f172a', borderColor: '#e2e8f0' }}
                            >
                              {quantitativoReportData.criticalAnalysis}
                            </p>
                            <span data-pdf-hide="true" className="absolute top-2 right-2 opacity-0 group-hover:opacity-100 transition-opacity bg-slate-800 text-white text-[10px] font-bold px-2 py-0.5 rounded-md shadow-xs flex items-center gap-1 print:hidden">
                              <Edit2 size={10} /> Clique para editar
                            </span>
                          </div>
                        )}
                      </div>

                      {/* Official Document Footer */}
                      <div className="pt-4 border-t border-slate-200 text-center text-[10px] font-bold space-y-0.5" style={{ color: '#64748b' }}>
                        <p>
                          CEO - Centro de Especialidades OdontolÃ³gicas.
                        </p>
                        <p>Fone: (88) 3614-3156 . Fax: (88) 3614-3245</p>
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {reportsTab === 'letterhead' && (
                <div className="space-y-6">
                  <div className="bg-white p-6 sm:p-8 rounded-3xl border border-slate-200/90 shadow-sm space-y-6">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-6 border-b border-slate-100">
                      <div>
                        <div className="flex items-center gap-2">
                          <span className="px-3 py-1 rounded-full text-xs font-bold bg-blue-50 text-blue-700 border border-blue-200">
                            PersonalizaÃ§Ã£o Institucional
                          </span>
                          {letterheadImage ? (
                            <span className="px-3 py-1 rounded-full text-xs font-bold bg-emerald-50 text-emerald-700 border border-emerald-200 flex items-center gap-1.5">
                              <span className="w-2 h-2 rounded-full bg-emerald-500" /> Timbrado Ativo
                            </span>
                          ) : (
                            <span className="px-3 py-1 rounded-full text-xs font-bold bg-amber-50 text-amber-700 border border-amber-200 flex items-center gap-1.5">
                              <span className="w-2 h-2 rounded-full bg-amber-500" /> Sem Timbrado Anexado
                            </span>
                          )}
                        </div>
                        <h3 className="text-xl sm:text-2xl font-black text-slate-900 mt-2">
                          Anexo de Papel Timbrado dos RelatÃ³rios
                        </h3>
                        <p className="text-xs sm:text-sm text-slate-500 font-medium mt-1 max-w-2xl">
                          Anexe a imagem oficial do papel timbrado do Ã³rgÃ£o ou instituiÃ§Ã£o (contendo cabeÃ§alho, logomarcas e rodapÃ©). A imagem anexada serÃ¡ inserida automaticamente no topo de <strong>todos os relatÃ³rios exportados em PDF</strong> (Estoque, CatÃ¡logo, SolicitaÃ§Ãµes, PCA, Termos de DoaÃ§Ã£o e Recibos).
                        </p>
                      </div>

                      {letterheadImage && (
                        <button
                          onClick={() => handleExportInventoryPDF()}
                          className="px-4 py-2.5 rounded-xl bg-slate-900 text-white font-extrabold text-xs flex items-center gap-2 hover:bg-slate-800 transition-all shadow-sm shrink-0"
                        >
                          <Download size={15} /> Testar ExportaÃ§Ã£o PDF
                        </button>
                      )}
                    </div>

                    <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
                      {/* Upload Zone */}
                      <div className="lg:col-span-5 space-y-4">
                        <label className="block text-xs font-extrabold text-slate-700 uppercase tracking-wider">
                          Selecione ou Arraste o Arquivo de Imagem
                        </label>
                        
                        <div className="relative group cursor-pointer">
                          <input 
                            type="file" 
                            accept="image/*" 
                            className="absolute inset-0 w-full h-full opacity-0 z-20 cursor-pointer"
                            onChange={(e) => {
                              const file = e.target.files?.[0];
                              if (file) handleLetterheadUpload(file);
                            }}
                          />
                          <div className={`p-8 rounded-2xl border-2 border-dashed transition-all text-center flex flex-col items-center justify-center gap-3 ${
                            letterheadImage 
                              ? 'border-blue-300 bg-blue-50/40 hover:bg-blue-50/80' 
                              : 'border-slate-300 bg-slate-50 hover:border-blue-400 hover:bg-slate-100/80'
                          }`}>
                            <div className={`p-4 rounded-2xl shadow-sm ${letterheadImage ? 'bg-blue-600 text-white' : 'bg-white text-slate-600 border border-slate-200'}`}>
                              <Upload size={28} />
                            </div>
                            <div>
                              <p className="font-extrabold text-sm text-slate-900">
                                {letterheadImage ? 'Clique para Substituir a Imagem' : 'Clique ou arraste o Papel Timbrado'}
                              </p>
                              <p className="text-xs text-slate-500 font-medium mt-1">
                                Formatos recomendados: PNG, JPG ou WEBP (Max: 5MB)
                              </p>
                            </div>
                            <span className="mt-2 px-4 py-1.5 rounded-xl bg-white border border-slate-200 text-slate-700 text-xs font-bold shadow-xs group-hover:border-blue-400">
                              {letterheadImage ? 'Escolher Novo Arquivo' : 'Selecionar do Computador'}
                            </span>
                          </div>
                        </div>

                        {letterheadImage && (
                          <div className="flex items-center justify-between p-4 bg-slate-50 rounded-2xl border border-slate-200/80">
                            <div className="flex items-center gap-3">
                              <div className="p-2 bg-emerald-100 text-emerald-700 rounded-xl">
                                <CheckCircle size={18} />
                              </div>
                              <div>
                                <p className="text-xs font-extrabold text-slate-900">Papel Timbrado Armazenado</p>
                                <p className="text-[10px] text-slate-500 font-medium">Sincronizado e pronto para emissÃ£o</p>
                              </div>
                            </div>
                            <button
                              onClick={handleRemoveLetterhead}
                              className="px-3 py-1.5 rounded-xl bg-rose-50 text-rose-600 hover:bg-rose-100 text-xs font-extrabold transition-colors flex items-center gap-1.5"
                              title="Remover papel timbrado"
                            >
                              <Trash2 size={14} /> Remover
                            </button>
                          </div>
                        )}

                        <div className="p-4 bg-blue-50/60 rounded-2xl border border-blue-100 space-y-2">
                          <p className="text-xs font-bold text-blue-900 flex items-center gap-1.5">
                            <ImageIcon size={16} className="text-blue-600" /> Dicas para melhor resultado
                          </p>
                          <ul className="text-xs text-blue-800/80 space-y-1 list-disc list-inside font-medium leading-relaxed">
                            <li>Utilize imagens em alta resoluÃ§Ã£o com fundo branco ou transparente.</li>
                            <li>O timbrado Ã© posicionado no cabeÃ§alho superior de cada pÃ¡gina gerada.</li>
                            <li>Sua alteraÃ§Ã£o Ã© salva imediatamente para todos os administradores.</li>
                          </ul>
                        </div>
                      </div>

                      {/* Live A4 Preview Simulation */}
                      <div className="lg:col-span-7 space-y-3">
                        <div className="flex items-center justify-between">
                          <label className="text-xs font-extrabold text-slate-700 uppercase tracking-wider">
                            PrÃ©-visualizaÃ§Ã£o da Folha A4 com Timbrado
                          </label>
                          <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest bg-slate-100 px-2.5 py-0.5 rounded-md">
                            ProporÃ§Ã£o A4
                          </span>
                        </div>

                        <div className="bg-slate-200/70 p-4 sm:p-6 rounded-3xl border border-slate-300/80 flex justify-center shadow-inner">
                          {/* A4 Sheet Container */}
                          <div className="bg-white w-full max-w-[480px] aspect-[1/1.414] rounded-lg shadow-xl border border-slate-300 p-4 sm:p-6 flex flex-col justify-between relative overflow-hidden">
                            {/* Top Letterhead Area */}
                            <div className="w-full h-16 sm:h-20 bg-slate-50 border border-dashed border-slate-200 rounded-md flex items-center justify-center overflow-hidden relative">
                              {letterheadImage ? (
                                <img 
                                  src={letterheadImage} 
                                  alt="Papel Timbrado" 
                                  className="w-full h-full object-contain"
                                />
                              ) : (
                                <div className="text-center p-2">
                                  <p className="text-xs font-extrabold text-slate-400 uppercase tracking-wider">
                                    CabeÃ§alho do Timbrado Oficial
                                  </p>
                                  <p className="text-[10px] text-slate-300">Nenhuma imagem anexada ainda</p>
                                </div>
                              )}
                            </div>

                            {/* Simulated Report Body Content */}
                            <div className="my-4 space-y-3 flex-1 opacity-70">
                              <div className="h-4 bg-slate-800 rounded-md w-3/4" />
                              <div className="h-2 bg-slate-200 rounded w-1/2" />
                              
                              <div className="space-y-1.5 pt-3">
                                <div className="h-6 bg-slate-100 rounded-md border border-slate-200 w-full flex items-center px-2">
                                  <div className="h-2 bg-slate-300 rounded w-1/4" />
                                </div>
                                <div className="h-5 bg-slate-50 rounded-md border border-slate-100 w-full flex items-center px-2">
                                  <div className="h-2 bg-slate-200 rounded w-1/3" />
                                </div>
                                <div className="h-5 bg-slate-50 rounded-md border border-slate-100 w-full flex items-center px-2">
                                  <div className="h-2 bg-slate-200 rounded w-1/2" />
                                </div>
                                <div className="h-5 bg-slate-50 rounded-md border border-slate-100 w-full flex items-center px-2">
                                  <div className="h-2 bg-slate-200 rounded w-2/3" />
                                </div>
                              </div>
                            </div>

                            {/* Footer Indicator */}
                            <div className="pt-2 border-t border-slate-100 flex justify-between items-center text-[8px] text-slate-400">
                              <span>RelatÃ³rio Oficial do Sistema</span>
                              <span>PÃ¡gina 1 de 1</span>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </motion.div>
          )}

          {activeTab === 'users' && isAdmin && (
            <motion.div 
              key="users"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="space-y-8"
            >
              <div className="flex justify-between items-center">
                <h3 className="text-2xl font-black">Gerenciamento de UsuÃ¡rios</h3>
                <button 
                  onClick={() => setIsRegistering(true)}
                  className="bg-[#1C1917] text-white px-6 py-3 rounded-2xl font-bold flex items-center gap-2 hover:bg-[#292524] transition-all shadow-lg"
                >
                  <Plus size={20} /> Novo UsuÃ¡rio
                </button>
              </div>

              {isRegistering && (
                <div className="bg-white p-8 rounded-[32px] border border-[#E7E5E4] shadow-sm max-w-2xl">
                  <div className="flex justify-between items-center mb-6">
                    <h4 className="text-lg font-bold">Cadastrar Novo UsuÃ¡rio</h4>
                    <button onClick={() => setIsRegistering(false)} className="text-[#A8A29E] hover:text-[#1C1917]">
                      <X size={20} />
                    </button>
                  </div>
                  <form onSubmit={handleRegister} className="space-y-4">
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div>
                        <label className="block text-[10px] font-black text-[#A8A29E] uppercase tracking-widest mb-1.5 ml-1">Nome Completo</label>
                        <input 
                          type="text" 
                          required
                          className="w-full px-4 py-3 bg-[#F5F5F4] border-none rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 font-bold text-sm"
                          placeholder="Nome do funcionÃ¡rio"
                          value={authName}
                          onChange={e => setAuthName(e.target.value)}
                        />
                      </div>
                      <div className="md:col-span-2">
                        <label className="block text-[10px] font-black text-[#A8A29E] uppercase tracking-widest mb-2 ml-1">Setores Autorizados</label>
                        <div className="flex flex-wrap gap-2 p-2 bg-[#F5F5F4] rounded-2xl border border-[#E7E5E4]/50">
                          {SECTORS.map(sector => {
                            const isSelected = authSectors.includes(sector);
                            return (
                              <button
                                key={sector}
                                type="button"
                                onClick={() => {
                                  if (isSelected) {
                                    setAuthSectors(authSectors.filter(s => s !== sector));
                                  } else {
                                    setAuthSectors([...authSectors, sector]);
                                  }
                                }}
                                className={`px-3 py-2 rounded-xl text-xs font-bold transition-all ${
                                  isSelected 
                                    ? 'bg-[#1C1917] text-white shadow-md' 
                                    : 'bg-white text-[#78716C] border border-[#E7E5E4] hover:bg-[#E7E5E4]'
                                }`}
                              >
                                {sector}
                              </button>
                            );
                          })}
                        </div>
                      </div>
                    </div>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div className="md:col-span-2">
                        <label className="block text-[10px] font-black text-[#A8A29E] uppercase tracking-widest mb-1.5 ml-1">E-mail</label>
                        <input 
                          type="email" 
                          required
                          className="w-full px-4 py-3 bg-[#F5F5F4] border-none rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 font-bold text-sm"
                          placeholder="email@empresa.com"
                          value={authEmail}
                          onChange={e => setAuthEmail(e.target.value)}
                        />
                      </div>
                    </div>
                    <button 
                      type="submit"
                      disabled={loginLoading}
                      className="w-full bg-[#1C1917] text-white py-4 rounded-2xl font-bold flex items-center justify-center gap-3 hover:bg-[#292524] transition-all shadow-xl active:scale-[0.98] disabled:opacity-50 mt-4"
                    >
                      {loginLoading ? (
                        <div className="animate-spin rounded-full h-5 w-5 border-b-2 border-white"></div>
                      ) : (
                        <><Save size={20} /> Salvar UsuÃ¡rio</>
                      )}
                    </button>
                  </form>
                </div>
              )}

              <div className="bg-white rounded-3xl border border-[#E7E5E4] shadow-sm overflow-hidden">
                <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse min-w-[700px]">
                  <thead>
                    <tr className="bg-[#FAFAF9] border-bottom border-[#E7E5E4]">
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Nome</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">E-mail</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Setor</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Papel</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">AÃ§Ãµes</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#E7E5E4]">
                    {usersList
                      .filter(u => u.email?.toLowerCase() !== 'gerlianemagalhaes79@gmail.com' && u.email?.toLowerCase() !== 'poli.almoxarifado@gmail.com')
                      .map(u => (
                      <tr key={u.id} className="hover:bg-[#FAFAF9] transition-all">
                        <td className="px-6 py-4 font-bold text-sm">{u.name}</td>
                        <td className="px-6 py-4 text-sm text-[#78716C]">{u.email}</td>
                        <td className="px-6 py-4">
                          <div className="flex flex-wrap gap-1 max-w-[200px]">
                            {u.allowedSectors && u.allowedSectors.length > 0 ? (
                              u.allowedSectors.map(s => (
                                <span 
                                  key={s}
                                  className="text-[9px] font-bold px-1.5 py-0.5 rounded-md" 
                                  style={{ 
                                    backgroundColor: `${SECTOR_COLORS[s || ''] || '#000000'}15`, 
                                    color: SECTOR_COLORS[s || ''] || '#000000' 
                                  }}
                                >
                                  {s}
                                </span>
                              ))
                            ) : (
                              <span className="text-xs font-bold px-2 py-1 rounded-lg" style={{ backgroundColor: `${SECTOR_COLORS[u.sector || ''] || '#000000'}20`, color: SECTOR_COLORS[u.sector || ''] || '#000000' }}>
                                {u.sector}
                              </span>
                            )}
                          </div>
                        </td>
                        <td className="px-6 py-4">
                          <span className={`text-[10px] font-black px-2 py-1 rounded-full uppercase tracking-widest ${u.role === 'ADMIN' ? 'bg-purple-100 text-purple-600' : 'bg-gray-100 text-gray-600'}`}>
                            {u.role}
                          </span>
                        </td>
                        <td className="px-6 py-4 text-right">
                          {u.email !== 'gerlianemagalhaes79@gmail.com' && (
                            <button 
                              onClick={() => setShowUserDeleteConfirm({ show: true, user: u })}
                              className="text-rose-600 hover:bg-rose-50 p-2 rounded-lg transition-all"
                            >
                              <Trash2 size={18} />
                            </button>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'trash' && (
            <motion.div 
              key="trash"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="space-y-8"
            >
              {/* Deleted Items */}
              <div className="space-y-4">
                <div className="flex items-center gap-2 px-2">
                  <Package className="text-[#78716C]" size={20} />
                  <h3 className="font-bold text-[#1C1917]">Itens ExcluÃ­dos</h3>
                </div>
                <div className="bg-white rounded-3xl border border-[#E7E5E4] shadow-sm overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse min-w-[600px]">
                    <thead>
                      <tr className="bg-[#FAFAF9] border-bottom border-[#E7E5E4]">
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Item</th>
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">ExcluÃ­do em</th>
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">ExcluÃ­do por</th>
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">AÃ§Ãµes</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E7E5E4]">
                      {items.filter(i => i.deletedAt).map(item => (
                        <tr key={item.id} className="hover:bg-[#FAFAF9] transition-all">
                          <td className="px-6 py-4">
                            <p className="font-bold text-sm">{item.name}</p>
                            <p className="text-xs text-[#A8A29E]">Lote: {item.batch_number}</p>
                          </td>
                          <td className="px-6 py-4 text-sm text-[#57534E]">
                            {item.deletedAt && new Date(item.deletedAt).toLocaleString('pt-BR')}
                          </td>
                          <td className="px-6 py-4 text-sm text-[#78716C]">
                            {item.deletedBy || '---'}
                          </td>
                          <td className="px-6 py-4 text-right">
                            <button 
                              onClick={async () => {
                                if (window.confirm('Deseja restaurar este item?')) {
                                  await updateDoc(doc(db, 'items', item.id), { 
                                    deletedAt: deleteField(),
                                    deletedBy: deleteField()
                                  });
                                  setToast({ show: true, message: 'Item restaurado!', type: 'success' });
                                }
                              }}
                              className="text-emerald-600 font-bold text-xs hover:underline"
                            >
                              Restaurar
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  </div>
                  {items.filter(i => i.deletedAt).length === 0 && (
                    <div className="p-12 text-center">
                      <p className="text-[#A8A29E] text-sm">Nenhum item na lixeira.</p>
                    </div>
                  )}
                </div>
              </div>

              {/* Deleted Requests */}
              <div className="space-y-4">
                <div className="flex items-center gap-2 px-2">
                  <FileText className="text-[#78716C]" size={20} />
                  <h3 className="font-bold text-[#1C1917]">SolicitaÃ§Ãµes ExcluÃ­das</h3>
                </div>
                <div className="bg-white rounded-3xl border border-[#E7E5E4] shadow-sm overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse min-w-[600px]">
                    <thead>
                      <tr className="bg-[#FAFAF9] border-bottom border-[#E7E5E4]">
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">SolicitaÃ§Ã£o</th>
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">ExcluÃ­do em</th>
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">ExcluÃ­do por</th>
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">AÃ§Ãµes</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E7E5E4]">
                      {requests.filter(r => r.deletedAt).map(req => (
                        <tr key={req.id} className="hover:bg-[#FAFAF9] transition-all">
                          <td className="px-6 py-4">
                            <p className="font-bold text-sm">#{req.id.slice(-5).toUpperCase()}</p>
                            <p className="text-xs text-[#A8A29E]">{req.sector}</p>
                          </td>
                          <td className="px-6 py-4 text-sm text-[#57534E]">
                            {req.deletedAt && new Date(req.deletedAt).toLocaleString('pt-BR')}
                          </td>
                          <td className="px-6 py-4 text-sm text-[#78716C]">
                            {req.deletedBy || '---'}
                          </td>
                          <td className="px-6 py-4 text-right">
                            <button 
                              onClick={async () => {
                                if (window.confirm('Deseja restaurar esta solicitaÃ§Ã£o?')) {
                                  await updateDoc(doc(db, 'requests', req.id), { 
                                    deletedAt: deleteField(),
                                    deletedBy: deleteField()
                                  });
                                  setToast({ show: true, message: 'SolicitaÃ§Ã£o restaurada!', type: 'success' });
                                }
                              }}
                              className="text-emerald-600 font-bold text-xs hover:underline"
                            >
                              Restaurar
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  </div>
                  {requests.filter(r => r.deletedAt).length === 0 && (
                    <div className="p-12 text-center">
                      <p className="text-[#A8A29E] text-sm">Nenhuma solicitaÃ§Ã£o na lixeira.</p>
                    </div>
                  )}
                </div>
              </div>

              {/* Deleted Transactions */}
              <div className="space-y-4">
                <div className="flex items-center gap-2 px-2">
                  <History className="text-[#78716C]" size={20} />
                  <h3 className="font-bold text-[#1C1917]">MovimentaÃ§Ãµes ExcluÃ­das</h3>
                </div>
                <div className="bg-white rounded-3xl border border-[#E7E5E4] shadow-sm overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse min-w-[500px]">
                    <thead>
                      <tr className="bg-[#FAFAF9] border-bottom border-[#E7E5E4]">
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">MovimentaÃ§Ã£o</th>
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">ExcluÃ­do em</th>
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">AÃ§Ãµes</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E7E5E4]">
                      {transactions.filter(t => t.deletedAt).map(trans => (
                        <tr key={trans.id} className="hover:bg-[#FAFAF9] transition-all">
                          <td className="px-6 py-4">
                            <p className="font-bold text-sm">{trans.item_name}</p>
                            <p className="text-xs text-[#A8A29E]">{trans.type === 'entry' ? 'Entrada' : 'SaÃ­da'} - {trans.quantity} un.</p>
                          </td>
                          <td className="px-6 py-4 text-sm text-[#57534E]">
                            {trans.deletedAt && new Date(trans.deletedAt).toLocaleString('pt-BR')}
                          </td>
                          <td className="px-6 py-4 text-right">
                            <button 
                              onClick={async () => {
                                if (window.confirm('Deseja restaurar esta movimentaÃ§Ã£o?')) {
                                  await updateDoc(doc(db, 'transactions', trans.id), { 
                                    deletedAt: deleteField()
                                  });
                                  setToast({ show: true, message: 'MovimentaÃ§Ã£o restaurada!', type: 'success' });
                                }
                              }}
                              className="text-emerald-600 font-bold text-xs hover:underline"
                            >
                              Restaurar
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  </div>
                  {transactions.filter(t => t.deletedAt).length === 0 && (
                    <div className="p-12 text-center">
                      <p className="text-[#A8A29E] text-sm">Nenhuma movimentaÃ§Ã£o na lixeira.</p>
                    </div>
                  )}
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'requests' && (
            <motion.div 
              key="requests"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="space-y-6"
            >
              <div className="bg-white rounded-3xl border border-[#E7E5E4] shadow-sm overflow-hidden">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-[#FAFAF9] border-bottom border-[#E7E5E4]">
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">NÂº / Data</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Setor</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Status</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Itens</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">AÃ§Ãµes</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#E7E5E4]">
                    {requests.filter(req => !req.deletedAt && !req.isReturn).map(req => (
                      <tr key={req.id} className="hover:bg-[#FAFAF9] transition-all">
                        <td className="px-6 py-4">
                          <p className="font-bold text-sm">#{req.id.slice(-5).toUpperCase()}</p>
                          <p className="text-xs text-[#A8A29E]">{new Date(req.date).toLocaleDateString('pt-BR')}</p>
                        </td>
                        <td className="px-6 py-4">
                          <span className="text-sm font-bold px-2 py-1 rounded-lg" style={{ backgroundColor: `${SECTOR_COLORS[req.sector]}20`, color: SECTOR_COLORS[req.sector] }}>
                            {req.sector}
                          </span>
                        </td>
                        <td className="px-6 py-4">
                          <span className={`text-[10px] font-black px-2 py-1 rounded-full uppercase tracking-widest border ${
                            req.status === 'PENDENTE' ? 'bg-amber-50 text-amber-600 border-amber-200' :
                            req.status === 'EM_SEPARACAO' ? 'bg-purple-50 text-purple-600 border-purple-200' :
                            req.status === 'APROVADO' ? 'bg-blue-50 text-blue-600 border-blue-200' :
                            req.status === 'ENTREGUE' ? 'bg-emerald-50 text-emerald-600 border-emerald-200' :
                            req.status === 'RECUSADO' ? 'bg-rose-50 text-rose-600 border-rose-200' :
                            req.status === 'DEVOLUCAO_PENDENTE' ? 'bg-amber-50 text-amber-700 border-amber-200' :
                            req.status === 'DEVOLUCAO_APROVADA' ? 'bg-emerald-50 text-emerald-700 border-emerald-200' :
                            req.status === 'DEVOLUCAO_RECUSADA' ? 'bg-rose-50 text-rose-700 border-rose-200' :
                            'bg-gray-50 text-gray-600 border-gray-200'
                          }`}>
                            {req.status === 'EM_SEPARACAO' ? 'EM SEPARAÃ‡ÃƒO' : 
                             req.status === 'DEVOLUCAO_PENDENTE' ? 'DEVOLUÃ‡ÃƒO PENDENTE' :
                             req.status === 'DEVOLUCAO_APROVADA' ? 'DEVOLUÃ‡ÃƒO APROVADA' :
                             req.status === 'DEVOLUCAO_RECUSADA' ? 'DEVOLUÃ‡ÃƒO RECUSADA' :
                             req.status}
                          </span>
                        </td>
                        <td className="px-6 py-4">
                          <p className="text-xs font-bold text-[#57534E]">
                            {allRequestItems.filter(ri => ri.request_id === req.id).length} itens solicitados
                          </p>
                        </td>
                        <td className="px-6 py-4 text-right">
                          <div className="flex justify-end items-center gap-2">
                            <button 
                              onClick={() => setShowRequestDetailModal({ show: true, request: req })}
                              className="bg-[#1C1917] text-white px-4 py-2 rounded-xl text-xs font-bold hover:bg-[#292524] transition-all"
                            >
                              Ver Detalhes
                            </button>
                            {isAdmin && req.status !== 'ENTREGUE' && (
                              <button 
                                onClick={() => handleDeleteRequest(req.id)}
                                className="p-2 text-rose-400 hover:bg-rose-50 hover:text-rose-600 rounded-xl transition-all"
                                title="Excluir"
                              >
                                <Trash2 size={18} />
                              </button>
                            )}
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {requests.filter(req => !req.deletedAt && !req.isReturn).length === 0 && (
                  <div className="p-20 text-center">
                    <FileText className="mx-auto text-[#E7E5E4] mb-4" size={48} />
                    <p className="text-[#78716C]">Nenhuma solicitaÃ§Ã£o encontrada.</p>
                  </div>
                )}
              </div>
            </motion.div>
          )}

          {activeTab === 'admin-devolutions' && (
            <motion.div 
              key="admin-devolutions"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="space-y-6"
            >
              <div className="bg-white rounded-3xl border border-[#E7E5E4] shadow-sm overflow-hidden">
                <div className="p-6 border-b border-[#E7E5E4]">
                  <h3 className="text-lg font-black">SolicitaÃ§Ãµes de DevoluÃ§Ã£o pendentes de aprovaÃ§Ã£o</h3>
                  <p className="text-xs text-[#78716C]">Visualize e aprove o retorno de materiais ao estoque.</p>
                </div>
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-[#FAFAF9] border-b border-[#E7E5E4]">
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">NÂº / Data</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Setor</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Status</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Motivo</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Itens</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">AÃ§Ãµes</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#E7E5E4]">
                    {requests.filter(req => !req.deletedAt && req.isReturn).map(req => (
                      <tr key={req.id} className="hover:bg-[#FAFAF9] transition-all">
                        <td className="px-6 py-4">
                          <p className="font-bold text-sm">#{req.id.slice(-5).toUpperCase()}</p>
                          <p className="text-xs text-[#A8A29E]">{new Date(req.date).toLocaleDateString('pt-BR')}</p>
                        </td>
                        <td className="px-6 py-4">
                          <span className="text-sm font-bold px-2 py-1 rounded-lg" style={{ backgroundColor: `${SECTOR_COLORS[req.sector]}20`, color: SECTOR_COLORS[req.sector] }}>
                            {req.sector}
                          </span>
                        </td>
                        <td className="px-6 py-4">
                          <span className={`text-[10px] font-black px-2 py-1 rounded-full uppercase tracking-widest border ${
                            req.status === 'DEVOLUCAO_PENDENTE' ? 'bg-amber-50 text-amber-700 border-amber-200' :
                            req.status === 'DEVOLUCAO_APROVADA' ? 'bg-emerald-50 text-emerald-700 border-emerald-200' :
                            req.status === 'DEVOLUCAO_RECUSADA' ? 'bg-rose-50 text-rose-700 border-rose-200' :
                            'bg-gray-50 text-gray-600 border-gray-200'
                          }`}>
                            {req.status === 'DEVOLUCAO_PENDENTE' ? 'PENDENTE' :
                             req.status === 'DEVOLUCAO_APROVADA' ? 'APROVADA' :
                             req.status === 'DEVOLUCAO_RECUSADA' ? 'RECUSADA' :
                             req.status}
                          </span>
                        </td>
                        <td className="px-6 py-4">
                          <span className="text-xs font-bold text-amber-800 bg-amber-50 px-2 py-1 rounded-lg">
                            {req.returnReason || 'NÃ£o especificado'}
                          </span>
                        </td>
                        <td className="px-6 py-4">
                          <p className="text-xs font-bold text-[#57534E]">
                            {allRequestItems.filter(ri => ri.request_id === req.id).length} itens a devolver
                          </p>
                        </td>
                        <td className="px-6 py-4 text-right">
                          <div className="flex justify-end items-center gap-2">
                            <button 
                              onClick={() => setShowRequestDetailModal({ show: true, request: req })}
                              className="bg-[#1C1917] text-white px-4 py-2 rounded-xl text-xs font-bold hover:bg-[#292524] transition-all"
                            >
                              Ver Detalhes e Aprovar
                            </button>
                            {isAdmin && (
                              <button 
                                onClick={() => handleDeleteRequest(req.id)}
                                className="p-2 text-rose-400 hover:bg-rose-50 hover:text-rose-600 rounded-xl transition-all"
                                title="Excluir"
                              >
                                <Trash2 size={18} />
                              </button>
                            )}
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {requests.filter(req => !req.deletedAt && req.isReturn).length === 0 && (
                  <div className="p-20 text-center">
                    <RotateCcw className="mx-auto text-[#E7E5E4] mb-4" size={48} />
                    <p className="text-[#78716C] font-bold">Nenhuma devoluÃ§Ã£o encontrada.</p>
                  </div>
                )}
              </div>
            </motion.div>
          )}

          {activeTab === 'my-requests' && (
            <motion.div 
              key="my-requests"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="space-y-6"
            >
              <div className="bg-white rounded-3xl border border-[#E7E5E4] shadow-sm overflow-hidden">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-[#FAFAF9] border-bottom border-[#E7E5E4]">
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">NÂº / Data</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Status</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Itens</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">AÃ§Ãµes</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#E7E5E4]">
                    {requests.filter(r => r.sector === selectedSector && !r.deletedAt).map(req => (
                      <tr key={req.id} className="hover:bg-[#FAFAF9] transition-all">
                        <td className="px-6 py-4">
                          <p className="font-bold text-sm">#{req.id.slice(-5).toUpperCase()}</p>
                          <p className="text-xs text-[#A8A29E]">{new Date(req.date).toLocaleDateString('pt-BR')}</p>
                        </td>
                        <td className="px-6 py-4">
                          <span className={`text-[10px] font-black px-2 py-1 rounded-full uppercase tracking-widest border ${
                            req.status === 'PENDENTE' ? 'bg-amber-50 text-amber-600 border-amber-200' :
                            req.status === 'EM_SEPARACAO' ? 'bg-purple-50 text-purple-600 border-purple-200' :
                            req.status === 'APROVADO' ? 'bg-blue-50 text-blue-600 border-blue-200' :
                            req.status === 'ENTREGUE' ? 'bg-emerald-50 text-emerald-600 border-emerald-200' :
                            req.status === 'RECUSADO' ? 'bg-rose-50 text-rose-600 border-rose-200' :
                            req.status === 'DEVOLUCAO_PENDENTE' ? 'bg-amber-50 text-amber-700 border-amber-200' :
                            req.status === 'DEVOLUCAO_APROVADA' ? 'bg-emerald-50 text-emerald-700 border-emerald-200' :
                            req.status === 'DEVOLUCAO_RECUSADA' ? 'bg-rose-50 text-rose-700 border-rose-200' :
                            'bg-gray-50 text-gray-600 border-gray-200'
                          }`}>
                            {req.status === 'EM_SEPARACAO' ? 'EM SEPARAÃ‡ÃƒO' : 
                             req.status === 'DEVOLUCAO_PENDENTE' ? 'DEVOLUÃ‡ÃƒO PENDENTE' :
                             req.status === 'DEVOLUCAO_APROVADA' ? 'DEVOLUÃ‡ÃƒO APROVADA' :
                             req.status === 'DEVOLUCAO_RECUSADA' ? 'DEVOLUÃ‡ÃƒO RECUSADA' :
                             req.status}
                          </span>
                        </td>
                        <td className="px-6 py-4">
                          <p className="text-xs font-bold text-[#57534E]">
                            {allRequestItems.filter(ri => ri.request_id === req.id).length} itens
                          </p>
                        </td>
                        <td className="px-6 py-4 text-right">
                          <div className="flex justify-end items-center gap-2">
                            <button 
                              onClick={() => setShowRequestDetailModal({ show: true, request: req })}
                              className="bg-[#1C1917] text-white px-4 py-2 rounded-xl text-xs font-bold hover:bg-[#292524] transition-all"
                            >
                              Ver Detalhes
                            </button>
                            {req.status === 'PENDENTE' && (
                              <>
                                <button 
                                  onClick={() => handleEditRequest(req)}
                                  className="p-2 text-blue-400 hover:bg-blue-50 hover:text-blue-600 rounded-xl transition-all"
                                  title="Editar SolicitaÃ§Ã£o"
                                >
                                  <Edit2 size={18} />
                                </button>
                                <button 
                                  onClick={() => handleDeleteRequest(req.id)}
                                  className="p-2 text-rose-400 hover:bg-rose-50 hover:text-rose-600 rounded-xl transition-all"
                                  title="Excluir SolicitaÃ§Ã£o"
                                >
                                  <Trash2 size={18} />
                                </button>
                              </>
                            )}
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {requests.filter(r => r.sector === selectedSector).length === 0 && (
                  <div className="p-20 text-center">
                    <FileText className="mx-auto text-[#E7E5E4] mb-4" size={48} />
                    <p className="text-[#78716C]">VocÃª ainda nÃ£o fez nenhuma solicitaÃ§Ã£o.</p>
                  </div>
                )}
              </div>
            </motion.div>
          )}

          {activeTab === 'leader-stats' && (
            <motion.div 
              key="leader-stats"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="space-y-8"
            >
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                {/* Top Requested */}
                <div className="bg-white p-8 rounded-[32px] border border-[#E7E5E4] shadow-sm">
                  <div className="flex items-center gap-3 mb-6">
                    <div className="bg-amber-100 p-2 rounded-xl">
                      <TrendingUp className="text-amber-600" size={24} />
                    </div>
                    <div>
                      <h3 className="text-xl font-black">Top 10 Mais Solicitados</h3>
                      <p className="text-xs text-[#78716C] font-medium uppercase tracking-wider">Baseado na quantidade solicitada</p>
                    </div>
                  </div>
                  
                  <div className="space-y-4">
                    {leaderStatistics.topRequested.map((item, index) => (
                      <div key={item.name} className="flex items-center gap-4 group">
                        <div className="w-8 h-8 flex items-center justify-center bg-[#F5F5F4] rounded-lg text-xs font-black text-[#78716C]">
                          {index + 1}
                        </div>
                        <div className="flex-1">
                          <p className="text-sm font-bold text-[#1C1917] line-clamp-1">{item.name}</p>
                          <div className="w-full h-1.5 bg-[#F5F5F4] rounded-full mt-1.5 overflow-hidden">
                            <motion.div 
                              initial={{ width: 0 }}
                              animate={{ width: `${leaderStatistics.topRequested[0]?.qty ? (item.qty / leaderStatistics.topRequested[0].qty) * 100 : 0}%` }}
                              className="h-full bg-amber-500 rounded-full"
                            />
                          </div>
                        </div>
                        <div className="text-right">
                          <p className="text-sm font-black text-[#1C1917]">{item.qty}</p>
                          <p className="text-[10px] font-bold text-[#A8A29E] uppercase">Unidades</p>
                        </div>
                      </div>
                    ))}
                    {leaderStatistics.topRequested.length === 0 && (
                      <p className="text-center py-8 text-[#A8A29E] text-sm">Sem dados suficientes.</p>
                    )}
                  </div>
                </div>

                {/* Top Delivered */}
                <div className="bg-white p-8 rounded-[32px] border border-[#E7E5E4] shadow-sm">
                  <div className="flex items-center gap-3 mb-6">
                    <div className="bg-emerald-100 p-2 rounded-xl">
                      <CheckCircle className="text-emerald-600" size={24} />
                    </div>
                    <div>
                      <h3 className="text-xl font-black">Top 10 Mais Entregues</h3>
                      <p className="text-xs text-[#78716C] font-medium uppercase tracking-wider">Baseado na quantidade aprovada e entregue</p>
                    </div>
                  </div>
                  
                  <div className="space-y-4">
                    {leaderStatistics.topDelivered.map((item, index) => (
                      <div key={item.name} className="flex items-center gap-4 group">
                        <div className="w-8 h-8 flex items-center justify-center bg-[#F5F5F4] rounded-lg text-xs font-black text-[#78716C]">
                          {index + 1}
                        </div>
                        <div className="flex-1">
                          <p className="text-sm font-bold text-[#1C1917] line-clamp-1">{item.name}</p>
                          <div className="w-full h-1.5 bg-[#F5F5F4] rounded-full mt-1.5 overflow-hidden">
                            <motion.div 
                              initial={{ width: 0 }}
                              animate={{ width: `${leaderStatistics.topDelivered[0]?.qty ? (item.qty / leaderStatistics.topDelivered[0].qty) * 100 : 0}%` }}
                              className="h-full bg-emerald-500 rounded-full"
                            />
                          </div>
                        </div>
                        <div className="text-right">
                          <p className="text-sm font-black text-[#1C1917]">{item.qty}</p>
                          <p className="text-[10px] font-bold text-[#A8A29E] uppercase">Unidades</p>
                        </div>
                      </div>
                    ))}
                    {leaderStatistics.topDelivered.length === 0 && (
                      <p className="text-center py-8 text-[#A8A29E] text-sm">Sem dados suficientes.</p>
                    )}
                  </div>
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'new-request' && (
            <motion.div 
              key="new-request"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="max-w-2xl mx-auto space-y-8"
            >
              <div className="bg-white p-8 rounded-[32px] border border-[#E7E5E4] shadow-sm">
                <div className="flex justify-between items-center mb-6">
                  <h3 className="text-2xl font-black">
                    {editingRequest ? 'Editar SolicitaÃ§Ã£o' : 'Nova SolicitaÃ§Ã£o'}
                  </h3>
                  {editingRequest && (
                    <button 
                      onClick={() => {
                        setEditingRequest(null);
                        setRequestBasket([]);
                        setRequestObservation('');
                        setActiveTab('my-requests');
                      }}
                      className="text-xs font-bold text-rose-600 hover:underline"
                    >
                      Cancelar EdiÃ§Ã£o
                    </button>
                  )}
                </div>
                <div className="space-y-6">
                  <div>
                    <label className="block text-xs font-bold text-[#A8A29E] uppercase tracking-widest mb-2">Setor Solicitante</label>
                    <input 
                      type="text" 
                      value={selectedSector || ''} 
                      disabled 
                      className="w-full px-4 py-3 bg-[#F5F5F4] border border-[#E7E5E4] rounded-2xl font-bold text-[#78716C]"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-[#A8A29E] uppercase tracking-widest mb-2">Adicionar Item</label>
                    <div className="relative">
                      <div className="relative">
                        <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-[#A8A29E]" size={18} />
                        <input 
                          type="text" 
                          placeholder="Digite o nome do material..."
                          className="w-full pl-12 pr-4 py-4 bg-white border border-[#E7E5E4] rounded-2xl focus:outline-none focus:ring-2 focus:ring-[#1C1917]/10 font-bold transition-all"
                          value={requestSearchTerm}
                          onChange={(e) => setRequestSearchTerm(e.target.value)}
                        />
                      </div>

                      {requestSearchTerm.length >= 2 && (
                        <motion.div 
                          initial={{ opacity: 0, y: -10 }}
                          animate={{ opacity: 1, y: 0 }}
                          className="absolute left-0 right-0 top-full mt-2 bg-white border border-[#E7E5E4] rounded-2xl shadow-xl z-50 max-h-60 overflow-y-auto overflow-x-hidden"
                        >
                          {(() => {
                            const allActiveGroups: Record<string, {name: string, category: string, id: string}> = {};
                            items.filter(i => !i.deletedAt && i.quantity > 0).forEach(i => {
                              if (!allActiveGroups[i.name]) {
                                allActiveGroups[i.name] = { name: i.name, category: i.category || 'Outros', id: i.id };
                              }
                            });

                            const filtered = Object.values(allActiveGroups)
                            .filter(group => normalizeString(group.name).includes(normalizeString(requestSearchTerm)))
                            .sort((a, b) => a.name.localeCompare(b.name))
                            .slice(0, 15);

                            if (filtered.length === 0) {
                              return (
                                <div className="p-8 text-center text-[#78716C]">
                                  <p className="text-sm font-medium">Nenhum material encontrado.</p>
                                </div>
                              );
                            }

                            return filtered.map(group => (
                              <button
                                key={group.name}
                                type="button"
                                onClick={() => {
                                  const existing = requestBasket.find(bi => bi.product_name === group.name);
                                  if (existing) {
                                    setRequestBasket(requestBasket.map(bi => bi.product_name === group.name ? { ...bi, quantity: bi.quantity + 1 } : bi));
                                  } else {
                                    setRequestBasket([...requestBasket, { product_id: group.id, product_name: group.name, quantity: 1 }]);
                                  }
                                  setRequestSearchTerm('');
                                }}
                                className="w-full px-6 py-4 text-left hover:bg-[#F5F5F4] transition-all flex items-center justify-between border-b border-[#F5F5F4] last:border-none"
                              >
                                <div>
                                  <p className="font-bold text-[#1C1917]">{group.name}</p>
                                  <p className="text-[10px] text-[#A8A29E] uppercase font-black tracking-widest">{group.category}</p>
                                </div>
                                <div className="flex items-center gap-2 text-emerald-600">
                                  <Plus size={16} />
                                  <span className="text-xs font-bold">Adicionar</span>
                                </div>
                              </button>
                            ));
                          })()}
                        </motion.div>
                      )}
                    </div>
                  </div>

                  {requestBasket.length > 0 && (
                    <div className="space-y-3">
                      <label className="block text-xs font-bold text-[#A8A29E] uppercase tracking-widest">Itens na Cesta</label>
                      {requestBasket.map(item => (
                        <div key={item.product_id} className="flex items-center justify-between p-4 bg-[#FAFAF9] rounded-2xl border border-[#E7E5E4]">
                          <p className="font-bold text-sm">{item.product_name}</p>
                          <div className="flex items-center gap-4">
                            <input 
                              type="number" 
                              min="1"
                              value={item.quantity}
                              onChange={(e) => setRequestBasket(requestBasket.map(bi => bi.product_id === item.product_id ? { ...bi, quantity: parseInt(e.target.value) || 1 } : bi))}
                              className="w-20 px-3 py-1 bg-white border border-[#E7E5E4] rounded-lg text-center font-bold text-sm"
                            />
                            <button 
                              onClick={() => setRequestBasket(requestBasket.filter(bi => bi.product_id !== item.product_id))}
                              className="text-rose-600 hover:bg-rose-50 p-2 rounded-lg transition-all"
                            >
                              <Trash2 size={18} />
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}

                  <div>
                    <label className="block text-xs font-bold text-[#A8A29E] uppercase tracking-widest mb-2">ObservaÃ§Ã£o (Opcional)</label>
                    <textarea 
                      value={requestObservation}
                      onChange={(e) => setRequestObservation(e.target.value)}
                      placeholder="Alguma observaÃ§Ã£o importante?"
                      className="w-full px-4 py-3 bg-white border border-[#E7E5E4] rounded-2xl focus:outline-none focus:ring-2 focus:ring-[#1C1917]/10 font-medium min-h-[100px]"
                    />
                  </div>

                  <button 
                    onClick={handleSubmitRequest}
                    disabled={isSubmittingRequest || requestBasket.length === 0}
                    className="w-full py-4 bg-[#1C1917] text-white rounded-2xl font-bold hover:bg-[#292524] transition-all shadow-lg disabled:opacity-50 disabled:cursor-not-allowed flex items-center justify-center gap-2"
                  >
                    {isSubmittingRequest ? 'Enviando...' : <><Save size={20} /> {editingRequest ? 'Salvar AlteraÃ§Ãµes' : 'Enviar SolicitaÃ§Ã£o'}</>}
                  </button>
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'devolution' && (
            <DevolutionTab
              selectedSector={selectedSector}
              devolutionSubTab={devolutionSubTab}
              setDevolutionSubTab={setDevolutionSubTab}
              requests={requests}
              allRequestItems={allRequestItems}
              items={items}
              isExpired={isExpired}
              isNearExpiry={isNearExpiry}
              setDevolutionBasket={setDevolutionBasket}
              setSelectedDevProduct={setSelectedDevProduct}
              setDevolutionReason={setDevolutionReason}
              setDevolutionObservation={setDevolutionObservation}
              setShowDevolutionModal={setShowDevolutionModal}
              setShowRequestDetailModal={setShowRequestDetailModal}
              showToast={showToast}
            />
          )}
        </AnimatePresence>
      </main>

      {/* Modals */}
      {showAddModal && (
        <div className="fixed inset-0 bg-black/40 backdrop-blur-sm z-50 flex items-center justify-center p-6">
          <motion.div 
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-white w-full max-w-5xl rounded-[40px] p-10 shadow-2xl max-h-[90vh] overflow-y-auto"
          >
            <div className="flex justify-between items-center mb-8">
              <div>
                <h3 className="text-3xl font-black text-[#1C1917]">Entrada de Materiais</h3>
                <p className="text-[#78716C] font-medium">Cadastre mÃºltiplos itens de uma vez</p>
              </div>
              <button 
                onClick={() => setShowAddModal(false)}
                className="p-2 hover:bg-[#F5F5F4] rounded-full transition-colors"
              >
                <X size={24} className="text-[#A8A29E]" />
              </button>
            </div>

            <form onSubmit={handleAddItem} className="space-y-8">
              {/* Common Fields Section */}
              <div className="bg-[#FAFAF9] p-8 rounded-[32px] border border-[#E7E5E4] grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
                <div className="lg:col-span-1">
                  <label className="block text-xs font-black text-[#78716C] uppercase tracking-widest mb-2">Fornecedor</label>
                  <input 
                    required
                    list="supplier-suggestions"
                    type="text" 
                    placeholder="Nome do fornecedor"
                    className="w-full px-4 py-3 bg-white border border-[#E7E5E4] rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 font-bold"
                    value={bulkEntry.supplier}
                    onChange={e => setBulkEntry({...bulkEntry, supplier: e.target.value.toUpperCase()})}
                  />
                </div>
                
                <div className="lg:col-span-1">
                  <label className="block text-xs font-black text-[#78716C] uppercase tracking-widest mb-2">Tipo de Item (Categoria)</label>
                  <div className="flex gap-2">
                    {showNewCategoryInput ? (
                      <div className="flex-1 flex gap-2">
                        <input 
                          type="text"
                          className="flex-1 px-4 py-3 bg-white border border-[#E7E5E4] rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 font-bold"
                          placeholder="Nova..."
                          value={newCategoryName}
                          onChange={e => setNewCategoryName(e.target.value)}
                          autoFocus
                        />
                        <button 
                          type="button"
                          onClick={() => {
                            if (newCategoryName.trim()) {
                              setCategories(prev => Array.from(new Set([...prev, newCategoryName.trim()])));
                              setBulkEntry({...bulkEntry, category: newCategoryName.trim()});
                              setNewCategoryName('');
                              setShowNewCategoryInput(false);
                            }
                          }}
                          className="bg-[#1C1917] text-white p-3 rounded-xl"
                        >
                          <Plus size={18} />
                        </button>
                      </div>
                    ) : (
                      <>
                        <select 
                          className="flex-1 px-4 py-3 bg-white border border-[#E7E5E4] rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 font-bold"
                          value={bulkEntry.category}
                          onChange={e => setBulkEntry({...bulkEntry, category: e.target.value})}
                        >
                          {categories.map(cat => (
                            <option key={cat} value={cat}>{cat}</option>
                          ))}
                        </select>
                        <button 
                          type="button"
                          onClick={() => setShowNewCategoryInput(true)}
                          className="bg-white text-[#1C1917] p-3 rounded-xl border border-[#E7E5E4] hover:bg-[#F5F5F4]"
                        >
                          <Plus size={18} />
                        </button>
                      </>
                    )}
                  </div>
                </div>

                <div className="lg:col-span-1">
                  <label className="block text-xs font-black text-[#78716C] uppercase tracking-widest mb-2">Origem</label>
                  <select 
                    className="w-full px-4 py-3 bg-white border border-[#E7E5E4] rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 font-bold"
                    value={bulkEntry.origin}
                    onChange={e => setBulkEntry({...bulkEntry, origin: e.target.value as any})}
                  >
                    <option value="contract">Contrato</option>
                    <option value="extra">Produto Extra</option>
                    <option value="donation">DoaÃ§Ã£o</option>
                  </select>
                </div>
              </div>

              {/* Items List Section */}
              <div className="space-y-4">
                <div className="flex justify-between items-center">
                  <h4 className="text-sm font-black text-[#1C1917] uppercase tracking-widest">Lista de Itens</h4>
                  <button 
                    type="button"
                    onClick={addBulkItemRow}
                    className="text-xs font-bold bg-emerald-50 text-emerald-600 px-4 py-2 rounded-xl border border-emerald-100 flex items-center gap-2 hover:bg-emerald-100 transition-all"
                  >
                    <Plus size={14} /> Adicionar Outro Item
                  </button>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full border-separate border-spacing-y-2">
                    <thead>
                      <tr className="text-left">
                        <th className="px-4 py-2 text-[10px] font-black text-[#A8A29E] uppercase tracking-widest min-w-[180px] md:min-w-[240px]">Nome do Item</th>
                        {bulkEntry.category === 'Medicamentos' && (
                          <th className="px-4 py-2 text-[10px] font-black text-[#A8A29E] uppercase tracking-widest w-36 min-w-[110px]">Tipo de Material</th>
                        )}
                        <th className="px-4 py-2 text-[10px] font-black text-[#A8A29E] uppercase tracking-widest w-36 min-w-[120px]">Unidade / Emb.</th>
                        <th className="px-4 py-2 text-[10px] font-black text-emerald-800 uppercase tracking-widest min-w-[110px] bg-emerald-100/70 rounded-t-xl text-center">Qtd. Entrada</th>
                        <th className="px-4 py-2 text-[10px] font-black text-amber-800 uppercase tracking-widest min-w-[100px] bg-amber-100/70 rounded-t-xl text-center">Estoque MÃ­n</th>
                        <th className="px-4 py-2 text-[10px] font-black text-[#A8A29E] uppercase tracking-widest w-24">Lote</th>
                        <th className="px-4 py-2 text-[10px] font-black text-[#A8A29E] uppercase tracking-widest w-40">Validade</th>
                        <th className="px-4 py-2 text-[10px] font-black text-[#A8A29E] uppercase tracking-widest w-28">PreÃ§o Un.</th>
                        <th className="px-4 py-2 text-[10px] font-black text-[#A8A29E] uppercase tracking-widest w-20"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {bulkEntry.items.map((item, index) => (
                        <tr key={item.id} className="group">
                          <td className="px-2 min-w-[180px] md:min-w-[240px]">
                            <input 
                              required
                              list="item-suggestions"
                              type="text"
                              placeholder="Nome do produto"
                              className="w-full px-3 py-2 bg-[#F5F5F4] border-none rounded-lg focus:ring-2 focus:ring-[#1C1917]/10 text-xs text-stone-900 font-bold"
                              value={item.name}
                              onChange={e => updateBulkItem(item.id, 'name', e.target.value)}
                            />
                            {/* Quick unit helpers for medications */}
                            {bulkEntry.category === 'Medicamentos' && (
                              <div className="mt-1.5 flex flex-col gap-1 bg-[#FAFAF9] p-2 rounded-lg border border-[#E7E5E4] max-w-[280px]">
                                <div className="flex flex-wrap gap-1 items-center">
                                  <span className="text-[8px] font-black text-[#78716C] uppercase tracking-wider mr-1">Unidades:</span>
                                  {['mg', 'mcg', 'UI', 'g', 'ml', '%'].map(unit => (
                                    <button
                                      key={unit}
                                      type="button"
                                      onClick={() => {
                                        let currentName = item.name.trim();
                                        if (currentName) {
                                          if (!currentName.endsWith(' ')) {
                                            currentName += ' ';
                                          }
                                          currentName += unit;
                                          updateBulkItem(item.id, 'name', currentName);
                                        }
                                      }}
                                      className="px-1.5 py-0.5 bg-stone-200 hover:bg-[#1C1917] hover:text-white text-stone-700 rounded text-[9px] font-bold transition-all uppercase"
                                    >
                                      +{unit}
                                    </button>
                                  ))}
                                </div>
                                <div className="flex flex-wrap gap-1 items-center">
                                  <span className="text-[8px] font-black text-[#78716C] uppercase tracking-wider mr-1">Dosagem:</span>
                                  {['500 mg', '1000 mg', '1000 UI', '5000 UI', '10.000 UI', '50.000 UI'].map(dose => (
                                    <button
                                      key={dose}
                                      type="button"
                                      onClick={() => {
                                        let currentName = item.name.trim();
                                        if (currentName) {
                                          if (!currentName.endsWith(' ')) {
                                            currentName += ' ';
                                          }
                                          currentName += dose;
                                          updateBulkItem(item.id, 'name', currentName);
                                        }
                                      }}
                                      className="px-1.5 py-0.5 bg-emerald-50 hover:bg-emerald-600 hover:text-white text-emerald-600 rounded text-[9px] font-bold transition-all uppercase"
                                    >
                                      +{dose}
                                    </button>
                                  ))}
                                </div>
                              </div>
                            )}
                          </td>
                          {bulkEntry.category === 'Medicamentos' && (
                            <td className="px-2 min-w-[110px]">
                              <select 
                                required
                                className="w-full px-3 py-2 bg-[#F5F5F4] border-none rounded-lg focus:ring-2 focus:ring-[#1C1917]/10 text-[11px] text-stone-900 font-bold"
                                value={item.medication_type || ''}
                                onChange={e => updateBulkItem(item.id, 'medication_type', e.target.value)}
                              >
                                <option value="">Selecione...</option>
                                <option value="PORTARIA 344">PORTARIA 344</option>
                                <option value="COMPRIMIDO">COMPRIMIDO</option>
                                <option value="AMPOLA">AMPOLA</option>
                                <option value="SOLUÃ‡ÃƒO">SOLUÃ‡ÃƒO</option>
                                <option value="SOLUÃ‡ÃƒO SPRAY">SOLUÃ‡ÃƒO SPRAY</option>
                                <option value="POMADA">POMADA</option>
                                <option value="GOTA">GOTA</option>
                                <option value="COLÃRIO">COLÃRIO</option>
                              </select>
                            </td>
                          )}
                          <td className="px-2 min-w-[140px]">
                            <select 
                              required
                              className="w-full px-3 py-2 bg-[#F5F5F4] border-none rounded-lg focus:ring-2 focus:ring-[#1C1917]/10 text-xs text-stone-900 font-bold"
                              value={
                                ['Unidade (UN)', 'Pacote (PCT)', 'Caixa (CX)', 'Frasco (FR)', 'Ampola (AMP)', 'Bisnaga (BSG)', 'Envelope (ENV)', 'GalÃ£o (GL)', 'Rolo (RL)', 'Par (PR)', 'Metro (M)', 'Quilo (KG)', 'Litro (L)', 'Resma'].includes(item.unit_measure || '')
                                  ? (item.unit_measure || 'Unidade (UN)')
                                  : 'Outro'
                              }
                              onChange={e => {
                                const val = e.target.value;
                                if (val === 'Outro') {
                                  updateBulkItem(item.id, 'unit_measure', '');
                                } else {
                                  updateBulkItem(item.id, 'unit_measure', val);
                                }
                              }}
                            >
                              <option value="Unidade (UN)">Unidade (UN)</option>
                              <option value="Pacote (PCT)">Pacote (PCT)</option>
                              <option value="Caixa (CX)">Caixa (CX)</option>
                              <option value="Frasco (FR)">Frasco (FR)</option>
                              <option value="Ampola (AMP)">Ampola (AMP)</option>
                              <option value="Bisnaga (BSG)">Bisnaga (BSG)</option>
                              <option value="Envelope (ENV)">Envelope (ENV)</option>
                              <option value="GalÃ£o (GL)">GalÃ£o (GL)</option>
                              <option value="Rolo (RL)">Rolo (RL)</option>
                              <option value="Par (PR)">Par (PR)</option>
                              <option value="Metro (M)">Metro (M)</option>
                              <option value="Quilo (KG)">Quilo (KG)</option>
                              <option value="Litro (L)">Litro (L)</option>
                              <option value="Resma">Resma</option>
                              <option value="Outro">Outro (digitar...)</option>
                            </select>
                            {!['Unidade (UN)', 'Pacote (PCT)', 'Caixa (CX)', 'Frasco (FR)', 'Ampola (AMP)', 'Bisnaga (BSG)', 'Envelope (ENV)', 'GalÃ£o (GL)', 'Rolo (RL)', 'Par (PR)', 'Metro (M)', 'Quilo (KG)', 'Litro (L)', 'Resma'].includes(item.unit_measure || '') && (
                              <input 
                                type="text"
                                placeholder="Especifique a embalagem..."
                                className="w-full mt-1 px-2.5 py-1 bg-white border border-stone-300 rounded-lg text-xs text-stone-900 font-bold focus:ring-2 focus:ring-[#1C1917]/10"
                                value={item.unit_measure || ''}
                                onChange={e => updateBulkItem(item.id, 'unit_measure', e.target.value)}
                              />
                            )}
                          </td>
                          <td className="px-2 min-w-[110px]">
                            <input 
                              required
                              type="number"
                              min="1"
                              placeholder="Qtd"
                              className="w-full px-3 py-2 bg-emerald-50 border-2 border-emerald-300 rounded-xl focus:ring-2 focus:ring-emerald-500 text-sm text-emerald-950 font-black shadow-sm text-center"
                              value={isNaN(item.initial_quantity) ? '' : item.initial_quantity}
                              onChange={e => updateBulkItem(item.id, 'initialxœì}KSÜXÖà~~Å5åøHºÉ$yÙ†¶ñGAÚíÛÐ@ÕtAØ"%@me*KRòhšÕ,æ|»YME/:¾/¢V³™-lÎ¹éJº/%‰«¬îÂ Ç}œ{î¹ç}>ü4ö†Y˜]ÍÎ“ “yÉiuÎ½h/^ÙYò’¼÷Þ“u2ò’4x3ÌZå×æænþ1\†ÇÏ2ßø<óI?òÒô½7^ÌŒ.ÛKdÛíÃÅnwty4cú¾‡£qFŒï’?Ã$ð-¯eW#Ãp<8’Ë»0È3]Û[£ÈëgqäÉ‹™w·¿mH°¸hŸŒ£ˆ H–Éè
àr|ÚöphíÕ.9Žh“ÿÃo/w»$‰ÇC?ðÛ—9‰ûãt=	‡§ð­ô‡h£K²à2k§ö/»½mŸÄÃ¬·ïFÎžöƒaf
Å—×!Ìà}+Ì‚AÀôA à` Û:©=1#!ñpëÌžBÓ€´d<ò½,ø~}z-±ŽBžÌÊm~5è~Ñ^Z™
–3ôÅåj„–oã,¸#Z~÷jþ·r$rƒ£S-.~·¸µ¸¶øôha‘£ãeÊÑ2ƒ&Úk]ŽŽˆ‰®Ø‡èpìeý³l+O»ä6kØõù±f¥kÃ?<—¿;‰‚K‚?Úý8"§Þ¨½hiÁóî!àl«Dˆ¦Þqø|©ÂôCxä¶/4ð!¸…‰•&˜r±³:e´<\\„ÃH‹šùœÖã‘×êÄØaÙ¤? í3wÅX©Ñ¦kAY¼žGÞqÕ0»OùÁ°ŒôÇI'íQÒSÄÖ°;Ú	ÄëŸýOÇñ¥æ´Y&gðG	&§‰wEOS¶ôœÇÜµÝ—XÓ…¼Ü|LÖwXrX™täe2­áî(¶ÜÓgOŸlÁèGAÒ÷Ò`fãM1Xß{¾€­Yñl"š…Ø- µ»‡£ùÙf7Î2Í‚°–îb£c¼;ßµ’œ¯ê/˜Èñ0Ì>Œ’°”YÈâþ´Žø¢E7öñU{œˆ¨j;;p<Î`E]ùö¶nÂòDaÿÓ‹ëÖ.?ÁŸŠr a2p 6	aU€]ê’³ø<HÖ‹{Oò{€ýôÎjWÆw†i˜…ñ°íE‘} wDÐñ6‚8â‘xœ%1‰¸lê»®Hþ0yñÙÃ½ÀÖÀöÞõ1€º7Ì’«E‹NO³3²AÉ¿ýiÙGæŠMq£†I0€%¨±_4À5~$qZÃz¯„ôÎÝð#Ç=:ëQà‚ úƒÄKÏ–š …;bX€zÇcž&º§::
ßÇþ•ê3x„\xý‘f˜Ïáð¢0ÍHè¿˜A4j§ãÓÓ Å%M•Ôðz3î°s’ÄƒÖ0¸ ûAÖb;fàZ!"hØ~ÍÍÍÑ[ø;ÞUo¡çñû"Ÿ‚«×øê8ÿØÊ•TAfÈçb›f
ÜXÂén*œ†?ƒ}þ>›aê8—4ŸH:•Yð¬ÞUkx”­QÖ~"ø˜Lürø]ïioµ·¢R+(˜bU(TdûgñÅ¦ï¿‹}/jxQª–·*CQ	Nô'È­ätf‰êò4œuN
ÞÍF”Tk±åûAä%ŠÑ‘	+¸Òññ Tª¢ª³>\:’æ},ñ“lºg°Ã4ð °´¶´ºT IÏ<?¾@rÍ+¸Õ¥.©sBÃž8¹’£e7(>ß÷ÎN†—ºˆöäÈ9Qø7àðp‘‡øî°½!p Sç5Pµç'q2ï=_ÄŠNéÝÒ§€™ü·ëÐv/Žo†ç0û8¹¢\æ j{.¼Ñ9Ö·»„2R^ÿÓÂ
üÿúI<BÖ*AEóßÚ‡OºGv˜ÚOJôy1‡2¾…À¯‡^ôâúš¤}/
ÖI·³6O¸þ"7åmçÃ0–Ò‹Òë‹Õ×¥iÂÄrÉià]¶/Ú?ÇËeÀËQû™À2@Sk*ë¤¢[ÇAvÃ2„Ç˜ÐfÎ–k¢xiw¨ùü%õÛM¨r¦Öž`”gb¯PÄ~ç(.÷@ÂƒC‚´vã$óæj(z¶\4'j¢YÇ>A>ëJ‡ï6Ÿm.­õŽd¾­ÐÖ(æø—Òæ¬UµÍTçNuõRÀŸ }¥Z 5ÛQSœGqÿS.9³õÃ]¤U¤ ‘ëBÙû"ôáGüX"ƒe¸ý 
ú!ŠîÙ÷"O«?©Îã4	}‚?PAœB{:dJ±·³ónŸr	¬˜)`@Õðx”UÀt\¦IpvYà#¢Ð¾µj1³ëp¼¬àñ²,[ÈÊðÆýÂÍj•äñu*uJtÒ/É¬î°Êùþu³åcº¼²UE>Øø­Ù›ê©ê8hdMì¿†1«c²ž'ûì8ÿ*ŒàiŒLœÆIè¥ÎX„ü¬½òŒ ÀO" ÝWmoœÅDlë%2ªêÇ$$*[bM&,HŸ/`Ì4üiØ9bt‹À‹7výÍ2hŽ8e¶	øŸ²¾Ê(ê°È¦átÕÀ—”+ä¬`qÕz÷\C#)¹ŽÆ èÌðò–Ý*ÀûÑ°aom+
Ãkƒ„žVM×nü‚TéX1¾Öa§ÓQ}žÀæþ`hø†pTNÜ·d'°·‚¤ÕG@ôÉ# ~;ó@´Ïnt´:µÅ¡D¶9ÅXYYé.oÁÖ£ûÅd\0˜¦@ýÔi3‘³)wô›.+€Qã:ì]?
z—#`JK0ÛÝ~UBsØTÀuÇüA5OÔ›@ƒí¶eQ}¥Ü_X$tuê±*JBñë êm ëÙÝÞÊ­æòîÂÆëãGté:xsŠ"ïªH¿pF/ðÆK®uÏÏŽªü”AVˆ±e”VJ¿êÕCv„ÙêÔ° Ánv/8SÊr-eÇ‹»ûÞí/pó¦‚ZUÙµü5-@ö©šëÅ5#9ÒpnÜDDçYÔ¹ÄçJSÙ"Ÿ²ÚN•#VA;y¤ÄÅ‰AWløáx §°v$TAô¨A'õ? ‰|Ê#2£õr;LGñðö—s`¨SZ¬€ÕÉ}Éxú°Œº±ê¦¡µÌöMc%L"žÚ¹¾útuy¥wDå¦BM ø†4‹—Åsv.êÅ «­Òù€åVçƒ?«ì| —„1‚Ÿî8ã7¾3ÿ/ñÌìËªÍ_×’ÁôÅ­%l`3²Úf<`þÎóö–¾™kÉø„{\+N–{¥R%·›Þ”Ñào³MðºØÁ¤…Î£ë¤îsIþþw2û~asöfŽ´ùsyGuLc´NZïbL?]§¦;RÕžZ‘AŸ¸ì¼ŠuyûÉ^Õ:G2³ñg
2¤>Ä#›~Èö«q«Ú]=ìÝ†jMmÓfÜw3é×e‰ôþ9»j½ó²3@åËð(Yé Ñ¢É±¦D\—ÈþÙc 0iŒP{Q¶é³[KÝªC=»ª+ø*ÜJ¡Ábm<ƒ68ƒ„žöNzÄokÑÞlÍwÓáØüÓhÌƒÁ+“‘òÊ5¡±tÐyë#¡é:b]–åòR]´——˜u…“·%à‡`1V»e–c-Á@Í û¥x·[Â<¾úápmê1éã½ë°^Gò™6heSþ6åo~Oþ~¢=©=ŠÔj%úänç‘ž6<P
1i­’ý`à½tÎ|Ô©ŠM)y°nyQç£ûâsp‘v_Õ’ª´)Þ…Ãý ?§•µðj™TW|šñÝu(GGô?Æÿ„C_x…>•Ë¬®Q]‹*llÇ¦¶f£º‚OÑÕì$û£wfx#=Ì™Ã#J™M½¢Ì“¡Üàéè(áïaÔ’ýŽ¬Î©‚Ç&ÓCçÏÎê¸™kU-ÚâÂ¡(Þ`ÇÔŸuâãVÃ4Íq ’Ð»yøJa6G¡YÙ¶Ð-¸„—ÊÏÇ0ö~ÌÜ[Ç)1éßþŒ»7&¨êÜþœ…ð‚“4LaÁ<5¿¯uÞzµùªiwó¼xû^¨ê»8ÏW/5âÚk¦øt°^ü¹b°ÌÓÖŒF÷	¹i |¦$ ðÒxØšE:1Ä³F¢#ÉÔûð3NZû½­ƒ½ýÃ®ÉÊ¥5 Õþ<°a"ƒÏ†iáÄŒ&±ù³§ìýüÓÚûMÜÈè¤rÂF?öúžoè	È‘fÖôö—A>‹+Û±wûÏÛ|­Èrû¡ß[¶ƒ´ï%Y°ð#~MíOÖz4¤NtûððM âÁ#Ül _)¾‚Ä÷šaÛ.~²°yî%¡ýË‡ŽflþÉd˜õ¢¹f¦#gc:†^vûsRöSy\ÔZZšRœh,æÎ+<	ûE³ûà>ú³‘Ä³:D?l˜ N®Ö±ý¦äúuãÏM«Ç…ªÚ

ñªPëÃZ.rGÅ˜1Å³PÀÇàaì”àÆ5?HY)ÓÛ!m²…Vè˜ù{‚~b"5‰’fG·ÿ:HÍí>Õ»¸Fï!ÅƒUÉpdÔlÕhÊäW]îß®téjIÄË×g	Äûrˆ¾ôƒ´ÞÇƒÀ¬Qäƒù¢£D²ÉÍþÐŸ€A+É®ˆàƒ‚iŠ¨2#óÙpÔŽ°lŒQ·¡ù} z–/	nÿ¨î–±bú©Ëç~ÊÇnÅíÏÝ‚OuÌ!òMßO‚4m˜¢Bšy#M0ßšÚÁFƒ¿ òn½ßýÓWˆ·Ýn§Ëþ[€ÿºí	SŸ<DìÅ¹êbŸoã±ë Áƒppœ ‹MZ;#ê5O÷§ÝÞë…Ý÷¯˜§”)æãH5=îŠ_‰Å9TÙÓ¯8iÉ‚Ú¹Ç‚ï¥gEôM!ˆKhÏeòŠ\nâŽÑCÒàE ôˆ4OLmeXçäIªT!#BaèÐÞˆ«QøÏÏ3ðN71j+ñÆ€KÆ¹Ë¥ÀO‹ü@gã˜%Ëz|¸ÒnA–OÂ(˜±¿ìõñøx1âD~çð…´2g¡ïC‡o
š×
¬š»¼#j°ÆyE>'ü;}Ù9ì™ÍÓìBã8~a3Ž—ûL `?¾ ˜$ä|¾Go´ŒJÀâbŸwâ!n™`èCCv}¥|Á™ð¶Œ“-Þ&°5ã(#^JÒ0ÇÝ48þ³™n{™÷ÃÞ[>—ïí›Á úd—…<8±8µíìéó¤ÆR,ë˜™¦‰±èÁ ÜYîØÆ~ï*H?{=´$ŠYœX¹›jº#B)Ä`
i°ÊžL-–ŽV
Ûmjþþ*ëƒx¹ó<ð†úÍ©ây5¡ÇEà'®öSapjÇÛ4é¿¨ÎðÆátˆàhØM‚ó0¸hv4äÓ§ÿÄÇú`Þ0óBëÆ±RËRáÁ1¼HeQauû¿£vˆsçazûøÁkÚz—ëdie¡»üäÒ“¯N·&¤6
oøwÛš»W/·Ém|m›á!uªšÌäjš‘²é*ÁÈešFß‚³ÇdÂq	£í<…3Ì˜L÷w £8©¡dC»ä?€¬}ÉÒk
mSŽX$áY©O®GàÔ¤]0ãàKrû3ìG¶;¾bRÒñvLAu–5×}öUD~¹ìÝ&:ñ©èÃ'ŠþÂuæTŸPÁ*-¾ò’ÁíÏýÐ#-îFÞ †`ÏÍl:Dœq7=–Â£QÔû¤HÈþÜ¿Ü_ô—:ÈKkÓ¿žˆjY>mHwqIÁÝE@÷<Ÿ@õ’â]vÓÏ)åKê3íðÓ™?ÑhÌðž{è$Ü„Ð

DÖ±ŠÝz ˆÞrÉìŠiÀÊŒ"L–IÏü7fÎûøN0&ÃXÄ£¡¦Q|“Aî(L‘ˆÄcâÑsßu¯ùLmXV¬Îs&ªó Q£m[én ù75!¼%¿(GÊ«sWf‡(®"5Dé«†ž%“ïš ’ùÕ;8_{é' âÖñ<ÁB —F-Q)¬D9rLƒK>„ú ;¡§ùd,Èå¼j¥Œ´¾Ðjõ­«ØmšùVÃó™´ü’Æ²P·Œ£!CäÃ`OÇ4Dláôêý(¢1¸
zfCŠ¥~©	¦&	óE·±ÇÉ9«‡›.¬‰&¬Û&»ªƒ‹ïé&|ÇÔZlÇ8è}ó)p~Aä¸Tý{±lïèû¾•·dýÊª-ÏÈü'«„*ÉG©È)Ds´-Úì­6“.Ž™]nººéWÞ°Å7ó·¼ËŠ½·¶¶f[~w&¬{æŠôÐOKÙïá°¥·5ºÚŽÆ÷¸Û`$_~;á¢½ø¤&N5ã C¸Qâ¶$®*¯¶gF/›Þò·@Xõhªÿö÷ßˆ­öÒ¨³Ë‘ØÞ;âkÃ¹Zžv²õ[úFHsO2&×jðªò\\ßX¯ñZÁA¤~Åe5¡üÅ½ü†Äïât£Aéµp3‰ã¼½² Íé¾!»>}£Éb8iíÈÛ.ù¯T†º†Å—?ƒ÷/éŸÉßzÇiœDÁI†ùJâQ{qa‰°#ˆf·¸¢7Ê0˜id*utÁHïWxÂYß”¬CöfK&¢Ý ýibˆ9Ïˆe·9+ô"êÏF‰A;bt‚*äv<Îãq…Ã vÂßÙ¬„WYE«$ÚìèÅ5‘wL¡ÔUù?Wÿ9×‚~®š1F:cÚÄe=oq²ÿ=¸ÚŽ/†fËò4
®˜ö®‡‡î,êkU ÷ø”WÁêºzÿQ`&ØíàÄG™«ÃÏÂa™$Iv†h•œ>ãN˜Šœúã™ÐžŒ:‚.ÍZ³øÜöxQw'xŽ¶ÚA)´ÁÆæ6™ÏáÄ£SÚWuåkù‚,ÑU¯â¬““]‰_\uþªšO¤×/.í©Ñ%Ixz†ÿâéÁr»b&œUu ‰EašàV—Cÿî®>×-¡¦¤Ü WTæ|?Ý™^JOðÃ#]²ùÊ[âÙ›8õõãÁ1~ôÆûøXÎƒøX“qöæ£«ÛU¨aËa½ `ŸºæµD‡sEÞúê+4wÖ‘tT'Y«åÍ“cJd=:Ù0^^lÅx[Ç¼Ø™S{Àï-Ø‹]§÷]“Z—-D¼|•’`:~Ó´x!^Elq!-æ'QÃzŒŸçÚyzŽ‰Ú‡®ç»hF©ØÃ³¡—"]ˆáÇ_oF+üzk°Ñ?)µCtf±œt’Ä‰ë9Â.†×î_¸.ƒ,Ëê‡yrM8pÖdæ‰ 4Ýµ¹€B­‹*äÄÂHßOÆ±ËÆ³Ëd	¤¼+’}{6AðÊ/^„²õ’fËV°Êmo¸øÖãe*½m6íTËzoH”ÜjÎÑ÷£7íð[ñ0.yêÇE»Ýv°(IpñùÏ_6
ílí‚¡ïïÀ[T×—Åñ"õ¢œ­»œYu±³ŠèÚ…DM©fxzÓ,Ópe:NqÕ¾˜Uû”v«”+uXßõGÀgeuÆcÌ3$ˆÙÁy¢¾—d¶ä1F}#ÊB?fÏh;ÕŒ¨¨TªÙN}2*wÖi”õ]¬¶0×Éâ·”cÀÇœg™eíï÷f
öæÀj¸:Î7ØD®EsiÿvÇ6B›·‚ÔÙüÖmÂÔuSAi2süÔ5|K²†É®[t«ÍÛ‘]wu¾)«”&MjÏ®†©íÙ5¡i¹NRVÙÙdcêÌºð½®,/Pž¤›ëÄ¹ÎÃN‚ì\ì$l½\Û91¿:£láo­jj“-®´“Û,»Xë<+ZVÚ»’jWPÈ²’‚jœl(Må8û šŠÅS¨ù .#÷º¿ð~ö†´,Øœ‚³xI>¶iB¸u`p&d>"ÓãAKjJˆËºæîÌ.ÈâãlpoÖ§³ä,sº]ë¸ò§¶¶Ý2V[oXà®aMu5©çõ×ªÉ±[×TÃ±Nh¶yÚ×z^JéF_ÌM[9q·E¢,;ûuÕÛ³s÷ÃÔ;ŽhÅMK!«Ë0›­ÛVðY×Íë5p*Šà•![ÖlˆA®sE:ŠÝë‡Iòê¹T¸ù”Ë_"®©ô’ð2P&3T®_<<	“—Økƒ©AÊÏ›­Ç×¥7sŠ|æêòÛ@³à^
í}«-Ï&äRMo¥oæ\(¾yÍkÁ÷.ûÑ8LÈ»R¸ZuªÅôŒuä”uÏhH…ƒÝÑ÷0®$Gcâ2ôÉ¸sM‘• ?ÆýÛÿ"£ØçtýñkGFhœ,pû¯„&®ò&â´S™IE=W~êìK3­ð™ÂÍk¤[3‡)`÷¾"K„ÖÂ’B¦­E¿õ†·ÿôhà®]Ãúg"É
£UG—("oK_ØãJtÞ.UY×¦ø¦AÖ~VGÝi¬.
,‘Ô2ãC]¾‹ûy˜”ú§aP´'‘#8XNéÐlU©lÇÿÑ<)ãNCÉG¾)©±Í!—³”æ¦õäíj N¡Lîvyaô­B®úØ>\[=¿8"Ñé:§˜ð;Åñè-ã+XÌC]8÷Ùj³Â¹*Âb²Ê©VµuMÅG çÇ`]ÁRHÕµÖ4áú5,*øX˜ë‡+àÌªõ®/iôÕ0ÅÃ]@‘ï½ð2Ö|°Nô1íÂ¬¹£Ýäö_—á N‰)UÓÛì&|Åø÷A?HSL¾í‰,Â^„Ø&*"•	ƒëÌ½«Ll ½™åê=Ú¸oQÉyp!÷NiŒ*õíJàþ(Ê¿ý¯!†„‡Øœ—Õs§+“J5Âä1dz÷œÔË•ÝßÆ´nÒîö+‹8Ž™ð™‰‘fÒ9M<?„¡´³¸“$´=TcQ*~zIbÍÈ·ôvqS¼ö¬|
¨ê€Â*ªj"«Ž#W‰è,ËÒ›Á(	À‡ï¡ƒ1r²”qb½•Üþ‚5™R`K;h¢BîÌ˜(zœŽÒAÇèäûY£á)É<ƒÝ:9<R²A•„XKÕúƒ‹R2,JÐ+ËSÞ‡Ž@üKInº›(ng3…ìQßi5R\¦X ž¯$æv ô5fL$@O³8ó¢¼Û,P
úºÖŒ})…ûˆF^²êsÂÞDÅ¯Ü.&(‘_šëèë¿±ŽáðÏ®}Éõäê=™«Íñx(j ¹'‰î=j@+ó‘%ž|| §êìJ×¡ÿ‡ºe¶ˆâ~“R§ ê‘øH´ŠÆq¿ÜœvP´©ÂÇ ÖZÅý@j°Ö"ðÒDsÿûSDØOŒVSÙßKÝ®¢ð~GÕ[*£“vé¬×¾¿…éèJÝ-
BÌþ\ëJ^øšv˜DpPRÿ]èØru éÛ|Ø{OšWÙßÏ®hNÊbÄ%a~¾Ì* ,ª5æ‹¯6„–A¤[’ó*©—€”àÇÞû­7Û;º÷dÀ‹6%ÍÒ³ìÈý+) N„BSÿšà,L<ˆ‰Àl½ è\ ¹À}8§Ÿ®$0ã_.PÞÝ»ý¿t»·çkÞ0?…àµ‡iÊÓR1_¸ñLç:b²J<×zÜS«æGêMÉèÍGMRÂª=A-n
©)ÓVUéÉ_³UW×'Ò˜èÜØbZé‘æÃÀL+m&üg*í^¼Àh§3üMÑÊËðo­ø²[<j…yq==NðS»[€¢|½,¶ÜbTÉ¯çZ:{^æ˜´XÅðÊ¶â&-Å¯8F‡¸kùÈ+XQg·L^’ç[^}O°ÅK]Ê¯[€¿97_5)¨˜<</oTÇµÅÎª9ê²îG«À®8¦IõJj à™ÇÃ>ü1³qíâ]«ó‡\Cç/á–*>¢Œ1/œ¿_2iõ|¼ä×ì8gv9´9Öé3¦Á d¦˜ŒNH‚ÛSc¶Ã
šV·ž%Û¬Ò_7—o‹0ÅÚb<&Á`ÎÞTrÓxdöÓ ­kÙdÎîË‰		ë-Q‰IæÊ#yª¨Ü	¾ìDf’×h§RŒUî†´xåv¸/I=®#0¥º5ì×"*]ásWøBaÅ”TÜ¨gmÎ£Wq2ú'õRØÿQ(€£ÖÏézòQ_«DG:è &vêy–´ÍÅ©i$“8óXó08GÝY£¢wîc"%³¦½á™ÆŒÂ7‡àa÷¨&ÿê{±yþÀÑ“O¢à|ÙÖåÑé§â;‘ÖØ\ÑÖØgcÝœË+b–uJ\(3Yƒ8ÂÇ×§­+‰©Û¨×C¯+Ä¨÷j× bfUÛ„'>d\Âý Ë€¶¦tY§é%òô×å%æ˜w¸¼D9¥²i©’ž”ã~¸ái­{~vô€M…u©ˆ ’ÌKÔz:NÐçãÿ¹õc²¦Ð¿×Ä–¢4_qrßEÞøu é§pBˆâÓ˜ä$Hzè”çÃ‘K2£;Fêl$ifã/í†%ÑfCõUÚr¢vóY]Ï­€´QÇm´N—î]/üŽ¼÷ÎÃSšxÇ)ùÝÂ+y$R&O}áY·$ÿªTc€2;«*–Œ/»5‹H£ŠuR98ÂA‚vªË²3F	a%Ìf1ª5šÇ*® -,L²§kçËjIg ­L€ÈÏ/­™TÂ¦'*ƒÍ]ÿ¶~â)ÐŽúë¾ÞÃœé†1ÓoqË4CÅ#¸Zµž„é&nzu”Þ†ÙÐã¹ŠFYG©:fî!a’—ØÈ5üÉÄè49B©PJc±û!’Ô–<‰!Õ+éP ø¡½tüV=?×½£pxòõ‘#:è¯‰Áx](Ñ~|œ6úS¾ënø}Rõ$ÂXuâµšøkÈ›®/¥Òªú6OóL—hµv$—óÎÐ—r­.ýKçZ„<äŠžÃ£,¤¦òG"5š!@êê],·âô©‡vµ=ôÙÊü&¦X‰Æxâ‰•…£X£LïŠ.JÆ7
<Ÿ:ä‘‡¢SyßH>EÚoy¾<o@â”2Åtl$>	aô!åÇŸÃãxxº±ÕÛüf¿Ï—ÄÃôö_I?„™îî¿ÛW¿õQx/ìÿP¼ ¸ôFqˆ³¿;¤H€ÅRþS©0	}tä‚ ’WØ§àE•Éb|ãxûc
ô”ìn¿šg%¸¨?KJhðð¸ê^À^©#3Ô• Îö-–÷Úæ’?²Z¼HVÍåª·’¤l¯{ŸR#áR¸éöÕ@U•øª	Cr9ŒŽMÛÅ1o„¾XO)£ÄŠ6~Álà¨ãe&I“C…$)Ì;ââ{{4Ž05…M½º›ÜþgûÇ0cÊ&¡ ¸»å·ÿô¢3ü+Í1A³½m–<×rpææél[AZ
âdÅˆódsÅÔ‰)I€¢ºCC9‡¼çÖŽªÑ¨†°áp$ÚuÄítÈjý6ß¾%Ë”X¦¨Cüc<ˆOƒaS,ryÃS$ê¦šŠ}ØÒáãð;tHd3Ä—^žÐºÄo1)O[ŒkyÉœadº´
yuÀ"¡ÊBî€ `-žÉ êŠŒB²•OkRNðÚ¤Ñi[M5X~V¬|ÃêÖÀ7S2R–¤S4çD3ôÞ|¼ÕAòÓè>Æ×¬”d•(Rñ<SDË0; œâ|fcóí»¿lî½yµ¹½3…Ú‰ÆÝ²Äw0¤µÅPpzÛ„9èÔ·ˆð7¸çí1ŠA4‹Â!p!q­ò¥Œs ©oÂêê†k¸îm˜Óûƒ½²Ý#½ýÝÞÖ›Í·o¶7·{ûdg{çýÁÎÛÛÿxýfksÿ¾‘qY c…Y%­=:ò»"æábgqõ±“rsq‚AjJåê8:eEó]ŒmˆŸ¥ïJØ™ƒ®1’2	Ìyõs¢+‡tŒÝy¿û{[ov„$sOÕeÍ‹øºžS¡öm8ð9Ç²2s¶Ä…Uç…¡äßHo¦X¿”úOc`*/¨\Ø„Ù’‹j&°kæàÏVë™þŠ…,‡Ë¬ÖÓ¸%ö€ü|ÏHÐ¤ÞëMüõÝæAo¨’a!ÏVùYøõU‹_«°• 1c\c´¾Gk]0·–¹y2ëûïÞ-\ÁEþøÇõÁÀ”„m²Z¢¶T$‘ø5Hé8þ•B×á(	Ÿâ—øš)DäÉuh#óã¾Gvé©ŠÞ3}O‹ªuÑ{µ^»®ÀNº\ŒA mD†®«IeðE.ƒ›·­XY©¦q ¬l±¦º3Þ•BæëŒþ¬ÊãXÔ×†CšˆŸÆÍÛEq¼è²ïP-UDåðÞŽiþZµîê>T‰Ò35¼öÇÇ°ÊÙ8Äü˜ê€Óô=T¦˜º Ä2±x§¤M(”Y¦Ü—æÅæ¬ÆÆÚ2ºT©„øíà[v+\¯±¸‚³•p€²$»‡ø– réK¦œ|L/eÐX\¦Ãl²õŒ=ÒgÖÔ“¶´Qº§š<œ!ë`ëÚ£˜…ÿÑp%wäëUï“:¹ÄLö‹€	Ôv¤JžZ“A	µN‚Ö¬D–—YZ..UàÖ
÷.—iêrØP´TxJäÊØU‹£ù$b CîÞÔöNråâ“f¢$íN©¡T¸àðø‘"  mDÔþ8‚?Ë'Á$:ß|@?Œ¢ØóË»p“ûy	îæÓ6ãpæØ¥ÆÂËdU9UŒÌ÷’aaKù<©º
±0E†è\([,¥¯ ¿"ÃÙøöŠY@ÓsèÝàVu6{FÍxeg¬ÊÏCÀy‚1qJØ­Ï®ôB£ÚùÝ÷¯çÉŸv_cEèý_£›÷Ï—²ôîû9‡>@vo‚*y‡˜§ÂI3Äë÷ƒÐ–¥¿+‘®™òô0. rþ3Ä1à‹1Ç¯é ”5¾Bß3eÙ€ë¾‰÷ äJ´©HDøŠÄ…bäDj(«ïMfØ*i"¿Iß$'‰¡ªøž-“j.2zYj(î:	…j.;4TÐßŸä0™Æÿ>Å×§¥¥úÒBDŒßˆ,‘#wiM¾IÎEÍ/î›xQ<ñ¢Ä*Ü¯p!,É˜ÅVyìÿ°?¹#iÁÃpHLŽ™á[˜šÍõ&å	áûñMò¸ÉCN>¡èÑ=¤8Ò»CßSˆ7•8ÊN¶Cú<eëvÝI
Qx¨Ý›(Rõfû&Œ|Fœ„¥ål…†/WKé•dé¶“PR¼?‰T2‰ßçý‰&fGRFc1¤à3Ë'R½yá¾¤„RÀâ7"žH…'ku)¿	(V%	Ä®"­w\†oBÊ„BJ…’Ý¯˜²ÂÅ”?=?ñ°€Æ»`8–‚¥§.‘Ð¸³oâÈ=ˆ#"âpRYD]ìàx±%Ì½ŽI³FÔÛoè§E–'”Ü§¸q¿2Æ7Áâ›`ÑH°¨H”JpQBfËr„¸ç$Dð—'” €ØP•rˆ$ð3‹yD¶X…/)(@üFÄŽÈÒB|œ„ŸÄ†ñ1Í¿'áàpˆÉâú"iU%…6ùí‰
÷/¬vÈ®7‚ƒñ 'ž/¥©hÑ,²¹2u!Gú&$L[H‡~xO(&ØPÁÈ+è.Fœx2Ã–—¢yÂÃ<!ÔƒŠ	X#~§ø ZÑYfb–0¯MïöÿÑ¸!JcrBgêÓt
÷&NDå
Å÷(VäýÊ…‹}¿IS“2ªXZHœÜpyƒÿU‘8Š»N2Gþú$RG}¤_Lú¨…K!»¹¢B x•Ài±Ðß~’NQbáKñ´´º_Hj	ó#¨Âßˆ “ïŠÒÊ|bÌBL4•C¿ÏX›û“e@|AÉnãXýŠCÛ²Lcx­&îÞUHB†·!©XÒâÓÆiþz/yAò²Ðøwú²sØ=Òç¤ÇëFá»s¼qÁF±½ÁšÚ¸ÑÔ/ÂëN2óåízvÒzÞH–è••Œ3¥È­Ñ¬'³B:¥›Ô•XÃ’ŠM—OÎÿº´¤¥gÂÖÇ™)%‰"³í±Ð‘0V-é×O_"VNh¹ZKøhOñ¸'IøW”‹p?„iÆªÃ¦d<¤µçQÔ9ÉËr)é³ŒtBX¦w01{x‚9-±$7¸ãÅOCZ”Ô@ð ŸA®b%NT¢“.É¢K&Ý:åRªíÖòï .û¼‚IÚÂ¢w)G1ÃÅ¬L.©1ó${êÚäA­@­6WjÙ(ð¬VÁÜ5éoQõ–ÿ&Æ	r€ZÊÓ¾½8ƒ3t«Q.MûŽ¢Z"•Ž	t¹SM¢ÝƒA¬Ur"»$MGä4sSBi´Í1e†‰"žOX•äíœ }…¸²ûà4N®èïÐøÂ‡³³¾z\ˆ¿íÂŠÅú™ó78/ÆŒ•º½¼–$Ñ)£5oyŠX]Œµ9Rx§e„â/pH=ª64ÆjsÎ¦ý€ÖBF zñÀö‹å†§ú˜œÊ-›Ó np”¡] ’”L8
-<Ñ²kêÉ¢ø­ž«)‘÷DæPoRpË8FÛÓ$cèdÈ3nÉZ¼Iv «‰ÉlÅÚ¯Öó«Q½ó6Œ‚¿QÙN±ŒMLÓXµ@+kè3›8ÒrA®¤¯0¦†yïä<îßþIÌ¥£ÏðöŸOÛÅ
‡È~2ÆzÀ>¤¼hL‘0˜'ã,„Oä4=©ˆ× \ö£q˜˜œ!»û¯$ìÓÔëô[ŒP%¾ˆopÃ¦ÐüS|ó„Ö‚:×ºieùÏWï­ñ™²»:@l %—’Á÷­t¾øì€ŠÏ­YN>ÄÃèÊá(À¼v÷uHäF‰)9ªOé€ãœàx(Ù=Ø	ÑãˆÞûeAÈ¡¦}?A•¸‰õ´ÚG#ÝÐ÷»ØÕªóln–-ô)ÈÏXë#(K‘©ÅyM	%Q¶œ·ò‰SI†æÚÐ¢­1xë€D;¢êñA\½ýGÜ0/=›F<ŒÕzÐ5žùIŒ[ªeg6;K®Cºú ¨¬×Š;‹	Á€ðJ®“€N‡3´w^„tJ .ÌPÄÁ¢ãv(º’k¿)74³±?Å	Væ ÛAÏãè<¤Ç»)©½Fu¥S´/{CØ›§^tæéÓµ?xaÔ.¡ùñÜìßi\³¬:šfÌg¿®™pŒdÞÕKÇú–‡ß-n-®->=šÙPi¯T•-V‘¬¨›–‘<üîÕ*üoåÈTFòó•ˆt/Ee.CåRƒ
vEA
m
x—rJEExPB˜ÜþLR‘Í	%T£“’®½T¸ˆéÑy%a‰)D˜º4¢/Qª ›bÎ—ÍÀQ{óÙæÒZïÈPÓÏ 8äâãÙ¹“„h•mí2|nñã9­áëy
"D_i™Tˆ¼Šó2)¡9Ç‚a<¬ÈýqºŽ•ßawHäxa±[ãp*ÎýÜ,yqÆã¤ŸoL•$SØ0±ŸKß´rC%mR±©•X<¢ú6Œ8f!N—cš’$ÈâN§ó|½®jéz<D[zN[:oÔJq¨j7Ñó§àêÅõG6ÿöãëôæãM”›øÏÔéœbšÀúÐuwÇp-…éý­2H	gF•Rjic3I@È/†´
P‘Ê·KB­í•ÕA3»êj‚ÊÃ{W“b_LZïÓ¶žý:64Û‹Í6ôAé›ûÚÐŸc;³?€íÌ
É«B«´ºýÚX(R¥*¬šjù.Z}ºº¼Ò;ª©ž(²IÜWïioµ‡Ÿ.5flyÃ~ ,¥Â:½Žl˜“K(*høaêGÿâ:Lñm4üýïäQùä£·Ê{oU^B‘}ƒÉn°•ÈÃyj¿;U´½9¢-—‚gåžÁÃ8Ã¡ÅO=eËz`?g]KKƒÖùÀ>w0ÂKh/J%¦#ŒŠ“}Én÷ƒ4EW D©[I#´Ž•úÁmýO[aÒ/Ì%l\¥r“·€Çl2µRï%Åé<¡ ¦ ~ß¤ÿ_¥ô_õI˜–€ù’üÖ¥)Œ=]Þd: Ù¹¦±€;Úa®æPú*Çº9Ú€D6árw:ÈöL^yŸÓšÈÜ~Yâ¿J}¢´¬) "Å!¼þ:„
Æ"àLqTM´â›û*r?N7±B‡‰Cø­¡¢ ©.ˆø¥$aÐ?7Ø?ßÔ¿"µØÜ¿z…A³½}Púæ¾ö6!~pXª‡ê›ä¾w9×<´]~_Ú3;öÔ$P€4×"„“$lùVå¥BÂ6í:ƒANø9	wfCE‚ÄÃ~Ó%üju	
§ìoJ…¬Tp
tS',±E/vË!ÃO»%YÍ˜nèš(9¯ŒÎµF¡51+MŒ~ï*Í	íF/»+“ŠÌlˆ¸8ôåŽ`#Œq9
NØN©“q t/2ñß•›Íô<ª ‹{QøTã˜èT"ÉÝõ?u”ŸžäRà„£ä"sÓÉyC±å3.õ¬²ªK'ÅH	5j2L5ˆ^/ÒôUáH’Ú‰Ee3À>cmHYäQ»†k£ä¶4nå¬S¤«¯1MôÈÜ,O B­v†‚ËrhƒÃ‹Ö0LZüÑÐšÓ:ñ›bµí¨G ‚¿Âu¾¡¨ˆÙ–vûE ïÄD ÇŠšO˜¼yAð—¹—9D9žÝÅ	1«†ìq2Ôæï(ËªÐz{€Ë¡“T5­ÀüésÒº.æs£K’e‚#ðrªyÜLEþýBtó}|îëô["”Ò®vÓiiÂ´õA2ÇÃÛðÍÝÇë¦Âm7`SÁoÒžÂ¿6èÏ¦ºŸÚÜ>|xßû>Ìlü0qm	)#žq~Í6“v•YÄŒ.èE’„$á§xæ	ül¯bîIÖ)$úÅ"åšO°£%ÙÀyK£¡À#Ž[›3Ë°œ28u<ÔFœÁn’3½Ëu²ãÃâèö_§a?ž'o=Øï^†#ˆ`‘Õ”«Ç9«cÍ{S"mÝ–xY¨W‘ mµ[!`üÑô(cMMÂ@¶gñ+lPñT)ªTìª‡†©"R¦¬‚m")ét±L~¶jc—š©cE0ØCPËR.AÊ¦Ÿýaä{Õ·%±_”2Aý‰åÔ(eKZVvÉ¤–ÍS†5
ÓŒN+qü,P“jo§„†*Û‚87ÑØÖÆzŸª[‰ºïEç_Lo+AÒ}MJ[EôDJÛÕß„Òö©Qi«J	±(r0ç>YÂÍÃµîùÙÑ”Õ¼jÍÕdI"ëJÞå<à=WðæáïVõ®*ç„Òë¾½R§_/Ç¬OOÃ+¥Äü©b´ßþûRNˆÏ«ÒU¥DhªÒ-ò™»«tåŒÍìæÓ¡ç%yrã«62¡d” œaã ™óŸ78†=ntýc¯,‰ÍÂþÄu1užGÓL•ÝHÉºy¦ËR€¹J˜â%	ßÊ)R´ét[V¹uª©Ë£näæxûŸ°Ãú },A#<œÇÅÎ;˜5Å’MS‘‘¥(0A<šaÐ÷DŸ„}äª‚ Eìi‡wpmÄlG[€*ÆZï¤zB#KBÂež’ÂÌÙOŒÓî»§¸Ð%qVQCŠ™¼€H¢O+ nN\WrZ×,•¦•Bµ.õZTk”uPègÚ\å+ª+y”uhâ©ÀßÌæDúÕ 5C7œVj	˜+¸5uÄ$ëà–:âkÔ^ðÓ8LC	=e@%Ð6žÉ`¥§+;TÁd…ªQCi',RÎ¹/}¤8añ’íœïo9‡ÎSNÓÄ( ÄÕ~j„rRÞØrLÙ4È	(ë!T¸¡Jé<‡£,É8I¬¶Xù•Aô+¨.·WQ3ç¯ëJ¹ì@fÓ~X-˜°T-˜×Q˜UÙ5võ,œa¬ÉÌ=¤XÉ ó)D5oöQrG•’(mzÙ;áN‰«díÅŠ%iUÅ‚©4¦J–MŸà¾’—ÊœÕž’©ÍQ0ôÒ"»Wá¦úB;w(±ãœÄInµ{Ã³aŸ„ å‡é<°u"Okq‚O>a-œÃR¼õ· £XR’z‘_Ž}¡vªtÃ&§¡2íŽ:¹´1»Ï½.oè€‘÷?}Ô«í76VÀs
tÌ‚J_1Ë©i³ÛŸi‰Á<›š,ñ lYŒ#…ÿ«3‘L¸­&ïÑxšâ?ªäå’à˜Ý÷AS¬òH¿Q«»P«2,ï@©Ðçk¡R{Ìä©˜~ª%ˆ×<Ù÷nAz€Y–%ùtîË’­·99ªñd Ls#³Î“4zfk&hÕÚ¡ïöŸ4V¹±Ø¦w£3œ„€)o*uu,ƒŠÓ4g±F¨žŽðœ'üut†QÌ[Œ˜W§@¦>E•'€U/>×hi„Ùc¦ä/³ž¢ñì ·ÐÃ¼Ú[;ï_½Ù{·¹ÇÖúÌ;#À„Dã8ÃÝf
ê!gŽvóiÓdvp±8ÎÄ\“ ÞÉi&?4.3yöw	¬RÅËe•RXHt¸½É”DßSJê ,º“×Cö\µÜw7V™½j
“UÙUwé(ò¾Vì]KõLÜ_ÐµÆ Ìõ¦I›×Í6¯Â ÞG·’–z+v@P´æ:Yü­- Z°„ðŒ§[Ÿê×š‹É>ïýeëíoöìÜdOL¹Ü(2e×)C»Ž5ý-Žž6{)éùçZêü»pÊª}pÿÎLLÄâfÔ3áþ}šL¹ý‹y<@Ç&8àöâx LÛ9š­“+Â›Šcîù&[„Ý³„÷ƒ¢Mê…_ç+œ¢œÝ¢¦ËjpŒšÀ5ª±s!Áe˜5é@íNu¸¶z~qD¢ÓuÎ]ÁoÌ»
IlÍ½j…¤ƒõŠ—•ì@Uö­*†N¾PÍ#g›y"uiÊ÷hëî&L£QíD¤#Ìsµèëæ¾óFTó½ïEiíÆIæ©ˆqg(Á³~÷ôÙÓÅ'[G„ÊÙ²T	¼>0$åP ®+
É³IùÀ)S—Iïî*5w}Ç»ñšwOo¦ÑÔSœY‰©ñ…é÷á<Âj‰°¨…vÚOâ(:öT~&9©5žõý«r¡iüùêÊ“bé%fç”c”*nŸÑŸ¥•€Îlªa,vˆŒñ¸åšl5³'Ð’‘T\ïíì¼Û§1S	.6hÊ¸YØEã©°}‰gç*R¬<sEÔ¡Ã3V3Î£`ðuš§Èò‰¶“«6™S5³—„ù³Êž½´ôšŽy“)¦Nµ¡tuÈJå<ë¬Ty \) ¯û¤Õ4³«Ð7¡#%}³¨”W„E‡ß­¬¬t—·Žt¹š×f—>¢µæT#¤§Ý’ö,ßò3oã>fg0•Âí`B_sMNM$`3#	U(R$&4ðü•,®¥‘}”)Æ { ²¡*9&­žf·?ŸQ#6å‘¥ =As¬VŠ`®±—ÓªÉƒõž jb…‡…mÂÿÖŽ
UÛ½2%«0¦z¬’‘
Y‹ÇYšªÎ1(’.ðh[ãÄÃÂ…é<é‚ä”y[hÔ»z—õ¾ÉC€§ËfX½¬+¢n?|©x[‘åyFY)Ž®’±zëi k™‰Q°íÒIZ Hë®?Òl8¼ŒµÇr|¾YØAˆ@Mp—¨–1OÏÐ!»ñ±÷ÎK0MïZÿ':¤tî¨‡ßm?Ù^Þ^<šÙø»¥\ÜôVàî@OÿÞa¾¤ƒ»‚½i]7¢¯•ÏOëÕ  ÝNRäKà)ü?³VšÙÑÓÊjs‰ô»Ò–-h5VTf:‹£
À'iÅÉ£ùú›Xv1F:3Ç—_>yáñ*÷D'ö£1Ðø|A[¤îü`(XJ çnõÉÄ¥ßØšNxˆEGÚ§jþ¢7“\	"ØË“‰½zTóù Žlƒ0<5H’¦mŽÝ\œ,šSßÊQ"‚¹§ÉŒ€—¡ô¶6Ä¡BNsa¶Ñxˆ€™Ù(v´å\Ð§×˜¢l¤‡ªÓGIøÅÒ½‚œ,ëKï	õc3.PòÓ±~r,Ý³zýô|Á™W¶ŒºÞ‘HíJTÖMä=­åePkõ„·…BUï.¤‹Ö+ÚÇà¼&ÕÖ±‹½Ãô^»Û¯U†7¨¢7Œ’ ÅðÌË•hù’Ÿ£ýixºQ¢ªˆ¯ì¶ž c‘ƒ˜d·¿dãˆ†Ý¦ñK¾cÏ¢Ñ™ª„7ãÐ0z„Csäöÿ¤ê1JÔ±ôámKÚî”+ÑÃKWò’â0¥è!‰b,F"eè{¨"$ÌÆÂñ'.
7ˆêÛ¾ßzµÎ)±Ÿˆ›Õ9\vPÖÍø\jþÂU`…Iæz9¼b
Ñ‹å¦êÆ *¬
(Õ¶‰/aN'½K¬Ä\8P€1OÊ;pž¨÷æ ·­©2¥ÙûÅ¸‡©º¶ë‚)‡KG2®˜Ë…MŽ¹Ö·¢ÇáJàËü·\™+fº.9Yä÷êÞn¾,ËŽ&-a¼”“^+Þ{MsðÊfÊ¦Ö¯ò-u²¨Ê«9x¾PsàO˜g ‹ßÜ2/Œ(²uð6ºhóˆÏo©°µYU–Uv½Á¿b«ü,‰²D²yM.qé„éKú’ÌâóèŒ%$ØÎãˆ¹AÓ$Mò³ýÂ§žÞ¨urN6ûœ¡œÙøÎ:R¿“B¿A«½Zñòº!mr=.È6`HËÒŒïà÷Ô‚àû
E­Y`
¾G>}r…ÊÍfyPêcj]Óý»NèÑ¡Ì|ú•5SZŸ ª?»¯2gnÎÔš?¹$†Ù~fc¿w°³GöwÞ¾Ùzs°ùþ §v©w­Òý²0}â¦À$Ä‰
;B”Tcái%Ï‰­KþoôðÌ}c¨W´0(é4giï²œ®…š`13/§dÄ˜HØŠ²ãú£&œ˜®%d¿ËR? ¥ÁÞ†l¸4ži·÷~»÷=jÑT/2åþæìO@h Þwë¨÷îÃ~owsosksGt6'£HònçO£»ÍÝ½7·ó®r}JI¹q÷Y½?Øë½þ!Ÿ\0±JîÜÙ^oë‡}iNêtfwîf»÷ãÎÛ`™>8!2ñšÔQS
_ÏM+ ŸNµ[òMÈµâû§‰wU¼Oÿz¢òî¾ù¨ÌMl%æ¦Ö{GØÛÿuû?wýR‰¢wGv›vBŠ*ˆLaéåÎŠÓë¬´àrgÅƒÉ;S¦¤Wè¯õnå¥{ÎL¹UÂë°¤ÿ{—ÆCUîêYŒÅ`IÓdrwd¯ò³Èz¿‹ÑŸ£"¸r<eUF9ù¯S Ò‘Ó©æ}¶u§ArÎb;¿ú¥š€èÉ²æƒW×åÍòó8—9AdVo7@‘JÞ?a¥™i°x¨´¿üðüA8Üi„$×Q¦,)|o—Vpase±‚ÑœˆuQ,¸t´æ‘©ª“v­«NB°^pxÅ÷
†o­~ò*Ï]×– 3µ/C“c[fê4!Pªô³*“à 5¶:C•£-P×°Ò:r‘w
6“³ËU™ßÉ•¾]Í“M¾¼³Å¹´ôÇ©§$Q›¸}C–%Qf˜hM¢?à*ñ-VÒU¥+÷E¿Ôò•ÃIWvhOëô÷$¾`F8®¹½“-ÎT˜	ÆO^XT_ë¼1Ç‹/™ÊqÁk´@,tåEïé/EYZ	*	EŸBŸyë?uàW$<üi‡`˜]}ðF£$>|ÇíÎ‘ø¡î_ysaÓ«õÏÈE˜Ñ0;"ç0E«9ÉÎÂ”[‚Dáy\‘,&˜U%¸„Wi"Ü„zŸÐÆ†c<YÕ~m9¶YKÒrh8?_ÜÛÌ`ÃåÊÜÊ£¹ôŒ_µ&–ß¢zÝü±i%8T²ÅùBÿÚÜ­v)¥+ Žâ¸æ]9Ñ«á5¡ÞË×›ÿ©ÿï¬ã§­
Y¤]ƒöyNWÉïÈ“.ü€S¶«lÀg€=Q_Å	ÐÂ eëäš%Még0]Ï:I©FýD ê:Ç‹?0,ùÀþz)^$7‡G°‡Gš^[{BŸä[IïWVT,riÑ+zñ=:tþ°#OGëÉ%ªâ±¹¹w¡Á·¬x&Õóúgl&g´
´;£qzÖ2{¯•WCšß¼ñ³b¹²œÊ˜¿—¿’ÿ¦uëÚí¶>V‹½ª´Ï¬|lõé`O]S66 £/Å÷¥-^£Ný ¥^ÞK ÈÇð²ö¨:ˆGa™RH«YoŒZ?ÔÆ8TÆ Jÿ&œ(pH{T™>/#”€í$8¢¼_+ÔÄ.±“e¹ìgÙé„“Ù+þ=ºèÌ¨ûºDéõ@¡[¯.¡þ}~^¿ñ×9ã Ï‹õÚñk$TpïV¥Óâåd-Q?¹óŠ¤k¶ú¯ˆw—D-•Ïì’² §b8²pÿ òC!>è\ª®5P3¶Â”Ê/¸úcv“ÎÈà¥Ì—8(±FfEÚ4vù²àßsêZ‚–`ÇXà»²âæ™=‚Ý¦¥Hâ¸H?™è]„…L‡fV€5àEIàùWlå€¢¾(ÚDÜFBk;º„÷°jjÝ®öfk‘Ÿß#ÝR©Ì-jHÏb<éà—Êa‡ws`yÍ[—£,ôK™‰.NçüÐ5½>ð.ÿìü²p¤à}Ã¦“s5°àeèv^vØƒYC¼´ö «.]#1?.nXc	ÓK¾Œ&ö·×Aì¥Ykæ šm°`‚ÚàÔ#½ý¥o ÂhæE´P=¼rGç!æ0š™'3áð$ž±¢ŽžÐ< iþPW'Ç÷tŽòTµ–¾cf†Öì{TeÁy 0žuûVRä´fß8zYZ(z.ž%ã`^pë6OD-Ë`MV1…™bªìAa·1E‘%7Ö€÷Ñ¦íÒð²+šòµçÁzé¶**Gã1_äú²0„X`ER}¥‚RÇLÔTÅ.úBÁSØÍ˜\uø¨±êPéR6¡¦Ú&PÏ«zG×#^ñeºjLX€ÌÕ¤µ3¢á ÑœFg­Årì£²»’†»³W8<&{ÑâŸL84]|]ÃØìb…+Ãê°:]Ð¯äÉË_Èk¥†’!6X™ Om]À©xÀÛ)Zá©(jVµ’½–Šb³òK2
9ÅÃ¤[`äð(Ä ¤˜òmÐm€9!`'¯xÈÙˆPµxL<?d	“B9ˆ?sÌÖ«Î™áš1C¯W„i–rdT& 6º©/vWÝóóVÐFK·mîÔj—D»¨u€RŠ*ªOã0óq=7rnH–‚5ïTtÊDNË3›­¼üÔH/ ôKÕ'Ã®L–â—àJj|^ß¼¤VJzœ"Ì-hœÅ£öâÂ)ñ*x£˜úJ‘ôÏÀq˜cÝ¥„8ºWJ‰_vƒ3Ð(”4œÓ Öä}ÁK±_£öâfœ+rl3&Ò‚ByvíÒ–´n`EÂm%ƒ™tS‰ò¦ï£¬Á–Rw˜iIsék·LAêBšÈj õQæêÔdÉÀèsŠK
*‚…æ	ü„uÒ‡ÎKáCEÐýÌð‘j“ü=UD$æ'×`NâÂÒþ< è,ô}àEy´ÑÒ*¥7†p£âÒ1×-[¡IŠ6©øñ†8J×‘e„Á?gF¹yr-óŠÔÅÔ¨°ßo6Èr­Õ%ŠHmî(©rû˜E)Æmð•¹2ãÉ‘-ñ„æ3œa³gwäÙ‡<¿	ªVvÆ˜®x–Aƒ’`0%~¸Ñ«SÄŠ1øQ¥ßÎñ_ƒ>ßµi«2]¢k¼Ä"œâ›Û!ò;sEÐYÏ©<ª/ÕwøÜœ±×4N²VË›'Ç7=ÚA'¢AX¨Þ^´uÌz5·CCÀ`Ç¯šÀ…8!€å¬|Êµ‰ªø©ôm¹ð1uÇ\ïƒá3ÇS0Äü(du­4Ø:õSácÊgƒšâ|ítÙuðââ”ÿi–bÁM‰Gì ŠK)ª¶ô0=¢ÌAÌYCqmž¿„8vÉµÄ©x‹}›?4â(Ç@¯št?BÅP$í¬ç<Fqdˆ˜>C&òæ’´˜áeÊð‘¿cð£–´E<©„-ÚLÆæ9K­U—H1¡¼3)iŒ¥Cc–9þÊnrG‘|¾¡ËdÑ“93¦©Ñ«Ö[ôÕŽçm:
øBž6o¡&J³œF‡ ²	>ß‰É;J‚D#rVLÏËU›TÙÅn
wqŽæºæ²žF‘î§®TÖqx^N5×©‘‚¸ô·
´!òF©FØËÎÏ×H[YR™@-çqœa.X‡PË¼»Ò1˜ÂÊê‹£0ë¦u¿Ï²³©7]®Aß[þœùŽ*¡Ï£ôVŽ*†²ë‡ÜÈœ.iP¦4Õ},‹¹uz"3=Uw–5ë”Hù6<ÆÜKžytŸpocj Äü‹Y8ŒÃƒG‰ZÜÕn¾çÙqì_ÉÃjú@M	ÿÅ¶×®§ê†€fß9Ê(
Ó­Žé]X [^ÔG^’§oò0šÚ[JÈI+fÉÊ€b"(­ÁDj—ÝÇB®Â@{PZÄÂÜ+€ÚŸK~ú&SƒœÐJÇÌ˜þxóßKbæ¼É€-9
s{»äÌPs¤#Ëõl¾ægñõÌ<ÒvÊªÓŽB¿Äµ8¤­4$LÃÆ}õs`kK‚Ts@MÖ›,r{[šíÂª*àM¾%‡Âf½ò@80ù&YÂ‰IÏ“Ìgƒ¯ ‡6®ÝŸ!swÖ¦´)ÊÆùÇ×Áx‚º!ôGŽÌ×†úT`]´nÿè€-iý’ç:Á“ÞÛ€™Y×P7.=­åÂSáQì™BŒ9i^õÎ47Ãt|ì.žë<-¢¶›ÔeÙÆ>š92Þÿöh`{9©ì¥ÉB¦8,ÍÔäá¼y]2Ó1ïmcæââ0¾˜1ê8Š‹Û@ÔDÚ¾¯ñªYEli€ÅÅð#/Iƒ7ÃZaW‹ÓeùBsL™ékU™ÀÂ)4)³4ˆášt:$œ'5P¬ÓqÞt´eC—1é°‰Œ_´—º¨[Æ´TKù¯Õ¡®ÁJ65cAùÒœÙ2ÙulÊÑ¼TÍR.o!jëZJàÈ‹~–GªØçZOÂ¨hŠ«).}=ù2ªºØå@äM‘Ê¥¦”'›"}CÕ¬ïz·CÌr„Ýítú¬‡SQñµÊ•¯>]]^ééµ9ùxKRÍKòeãuØoòýR€ÑÍGª,±f”*Ko¢ƒÒal*`ÕIáô3'©&-8ÎUa×T'7iCnæž™;¨>¢YWÆÄðËf.˜'DæxÝmUÁmúi#æêAxqHiu‹‹Š0óÔ^€Uw³”[¹©Ü3Ñ&þ*{—ˆm].-\¯Þ¼øs	­geæ`Ë-«íÞ¯¢øÂÈŸZWGg>dsèóH;è¬‰qú6ã© ÇKp›ìkÇÄÉKz^Û¨·ÙÂb#[aÒ*›m/!ß{á%­¡ÆåL]7›Í
âÿ  ÿÿì]ëRGþ¿O1VR‹Ø @`0K9¤ˆ+ÞJìlHR[KQö 6’F;#ÅÃøïVåWÛsú6}9}!'»“J"¤™î™éîsésÎ÷‘Wßmîü.§KÅ›ž+„îÅÆýV&æïâÃëÛs÷;=ªÎÅ³‹)¹å®Åpõ~]}ÕçÜMeéù_Wå ¨ëÑômÓzKe÷ÐPv„&#Õe^?ÉvI
Q‹*GÏï¡K­ñWM÷11šKªòŠ.èºÃ¼˜ï¤!ýÉór"u¶­/’T{õFø¼·æÿŽÚ&h¶[Ø›ŠÐ+Æýr¯l?û”ð*©:â«‹õhÍ;?cDyíÔÊKä,EûUb7˜G¹Ÿ ´¥?WZÑítQƒ›¼Ô+al ¦‘­÷Î„_–•&9QÓ.e›X`[›Ìº~RNßŒª	,£#Qû\pk»
ÚÚË¨K©zÐºÄóS§ðNhÉ0>Jµ»BÈÝµÎZ`ÈG©{éC§ÂUOñ+Ü}ˆ…•*™Jƒ/bËt0EikØ¹?f©åðS–ÔÒb{i	£¼ÌÝ¦åŽò·U^ŸÛÃ|…$s‹Ž3-¬ôß©ÌZãªF¸1wÑÂ=P”K-•xŽÊ_777vIR¥`
¹¢Tzˆˆ˜XìhWé›Â´¤L5¬K»ª%½VJÒ.¥P-ùé1ù+ùµ±·©×C9äLøÀì¹÷2YÓ¨(™¶J¦L&0ïúÉ™±û&xö9—	¶=ç‘4¡ÃÇæC•“UH$OŠE—»ÙL‹žJ
uþÂ7Z½C¢¨€;G"7J^£}¹g³¾:é~ß³EŒ¯õ‹"ÇQv^h"ó6’4kNÍ™9[(Ls"a[b¤onô7ÎgUkå½ “É˜­-’-ÉC…C4HðxQ>®ñxçà8Vº_…©­
×#KMT„G wŽDö‚NJ˜°Ådßá’8ò8ð¾€-:–F,b›ALÆÞ`º½ i7(0yu\^Ì£ÓN›H‚äýG›LÜº5úìE=Ò^”ŽÃ¬…è4ßž«¾ÇUŸW£é=EÝ›0±„%}*RžŽfÑÑe´
OÊI™½¹˜²ì,ç L|R~–ŒÊÏ^ïn"ÑïK		%+¿1)yÊà¡jÝt 1]áÿr$¬e'ÍŠ1ü9ž”à=Þ`]^v8»ýKHr¶WÁ®¶95Hs½M¡#dËÅDÝ²šæ¼YÐ9 &áŽ>•C‡,E¿aÂupï6éŸ•Õ${6*ÆÃz_"½üYCA,‹ØJ°˜ÙúÙd¸oµ¹¬·1FNè[ñ–ÌñlßxHvO6‰ÌÇéüÅ—"â`à‹òD
ÚÐBòJÅFqÀÊïC‰zôž-×3ÁÎÈ²%=Ë.’¶›Ì Q¯[VF’ú%Tbc Ðæ¤ËPÎØ€ŽœúPÿóñ?7¡™ï‹)BqtÄ‡6—JìYuûËXÌÐˆýU‹æ>ÈÌï7(oà®ˆ/[4yTUe6µ¤bç€üºE³¼à½sÀÿï¿–[F÷AJS–nd’õ´…ìð¦Ëj€')‚eä%#°%üRçèj¿A«©Ë³
á‘˜šœaˆc8B}WzÐVþPÂJØw1q•b~ùTúáp˜1ˆòc¶VÀÈF+™
5øP=66²þzö¬Àd>f~fƒ‹ª‚!¿ËE Mg5¦ÁÚþ(Kiæ'°4Y‰Jë‚u8Ïl—i9'tGÜ=
Ûå7œ L	ëZ_±ágy5¹}?å+¼øËhCûuu•ìÕ±~v«½àÝmÉw§°áUµÆß™:÷›âßÆ¯hB,Òa@'›Ï
÷^ùéRTÍËuöã!vìàu[â™ãü&0|B·!ê¶¾Êg.ô‹š+¿¥HÈôsž6ªý$a‰›Ÿdµ\ÃACÈX:±%ãÃNF þ®ÖÏóºkDŸVi0dlä®x™Ü¼å·"[àÊ3@›fpN,ôá ¨Mà"|¹ÞøêÉ$“%À@ÌŸÀÙÍ”²p‘½ðÇ¤:Úx¹Äôø„ž6t‡IM«M·œzîÙß¸Á‹€²PŒ2Ra˜`DÍÓð"dŽTãÎ>0!fîh‚»>cÊÉ~6#—ÝÌ}rçav5ë{¶®¥a«Alƒ G?+®f#T\•¦(Pq—ìòØZÕÐ¦£úˆ_Óeå9îÍ<\GÈÿ30N2n]s:&ŒËKÒG
,|øŒ´@ÿ¡xçJò¸t×†Ã'º×àžUå„=ígû`Ñ•ã\r%Ñ­ûußâõso¹ºñt'¯Ùhô>¾fW`AôkŸ€cMœO‰ã­ˆ3Cçñ7õZ6ËêŸ³YNÁ¿tˆZŽ7^¬Üœ’ýkïš_"áÑF!p«ÚÈpÄnWØŸÀ!£–½µ¡Ô=š1t a ØEš¢¤ÕZêÕ^ýÆ'€ï:Z›èÁ¾&|Ÿæ5~¨8¿D}AŸßŒùP›\#[þáA¨6Jæ“¢DÝÅK!|l	oŠÛŠÏÀ‹9ä-j‚2ãrõ‹ÑÛó1F@²Ïóé”7‰]‘¬Ñùò£Ü¨Y	èÚÑù@T¢	ÝGB,PÏ•JwÚlˆî©bñf¿ükßFI7‚d7%µ"ì¦òPñ«LlÆÕZAØVÉþÉ…¤÷‹ÝŠ:Uç©gRÌ®z[¬äèýf¬–œ"ÁWèŸ¨7¡IŒcKU<Ëüm1|ß¸ÈWú;
•Mß À‘à;ê•¥ÄqøñÜÜÁ<ŸÀhOÔ”Üã›©ì=äCV^"1j†Ý¹À=¯²[‰czÃïgåœÅfÊlœYT7ÆFG¾}ß»!v†ÀÂvôpÉŸ‘Qõ“GâÑÛü FíhøÚ¯j›CPÖ‡œºç¹ )!>´xKÖërRtÏ°Á3C[±&ºãø; 3ÏÖ-þã¬ îMÄzYç¦iõ†Ì/ÁX+KÙÖV¡“k¢Ä´Ÿ_>ñyXWå©„‘tä3$S‡¢«Wµþï\ëƒ‘‚»H.S±aní’ë[3_ÜÙ=¸™ËP‘síh“XH$pð˜øg›Ìñö•n9¸V<¨²¶ŒÞ½â¤¨óòK†5‹gÛxÚÏ¿YYÕºþû|˜"þUê“™£#íõÄªô ähNIAñLÀzÅ#ÂDÕ”&qPDB6ýà’Æ‘‹¹DlÝXoäF˜·Ô¾–›íú"K»¾aû²§@Úõ’þk±«>°æñ®ÂT7]¸lW'ëëë¶*Y“Ãvš„7B’a	³ÏOiåmAw½¶Ò¢³°ŠBñã;‹¬+ü°<[}ð	¨+:âJ¶BÉâTWÕB9½}P±*1)	•Åµ=’¦3Ç^iò51ûþ\TuYõ¦åï½¼DÎq:„˜ˆÁ²ß`½µhÊŠÕêF?Cu%Ž¸’ -æ¾Bx;»}Ÿª¤'fð+V¯®X=4‰“~Ä‘‹ùYµdnö×ky[õþä¯×:žçÓa^›(êS±CêÙ |="Þ<Îu´+wÀ¿f
sÄã·p“[-V3õöîûsÍ€ºÞqs•xˆž9@[d5ÃÆ‡,Ô3x.¤2ýÊý¦8Û«=WwÞY'2(¤R„•ñ5×îþA%Q;×§q¤àA”bø¹eÚe<$ñÍ¤d=¤§añ#°‘bæuz=5‰„KÍ"æWëj¹×åñãÚÞ…È"#Úó ›;3¿BÎvm·×:ƒÆ¡ûhÂ]aá*Þóugëð¿9}ù_Q'ó®Á(y:ªgåôöçð]®µˆW`£ŒÝJt ÉÞ/Ïé/õ{Ä/IóGZ¸,âàŠ€ÕÛ+ä/2C[óÿŠ€"’‘×¡uêoOô£¦I"TŽ &‹®RÉ”›nÓ}0Ýû²»‹{a¦ÿEÝîz=æÝW[Y=Ù<µÚxeýØ©Ò“]G0Íáø^²…»2ãqÍ–î”Ñj/@.ÌnÒoÿ¥ÀI«4I±êeuº…¾ÑÏ3\Ð7²Ê«öÜHyóû¢,Ü’æ;Â=Iv¯ØpKrZ^VùlelÑš€A×âðJð0‚‚Ç(÷ÑŠR_[³ˆ³Ÿ_¡1ÿåäNù¦I¼hfsÆùµ£µ~[|Q7’%)Œ–fÉ{Ø‡-–"¾fšñœÿdš³šsÎ?þm¸_;&T<|šVÓjðƒ5=®¦Šµ|L
ž[¿ñ?”NkÂh6dzc×G™‡	43›`S£Á!¤aiì›³I,üÅ˜¾.ÑfÈŽw4Þ–aî¸>ýg\P¡(—âÀ¦†w2=´õg´mIåàˆ¨c¸C	í$fÄ£`tðÄ¶£(™DJÅ”¯1÷È;gû¼;››f}5Ê [MB±ky|uûþj=ã÷ÏÊø\üÁüüªÚ9÷5^†ð®‹ÉÈ~{HÌXuöÓéÚ\Úqidß ·$Ï>t”xËJOÂ^!žË¯FGÞ#eŸ¦A±‹FN@džº1÷ÌÜJk’òD7IÑ–$txºZ¨ó¾^öÝ¿õá[kÖSp?¬Oî‡Åc)!o{£z;‹ïöÈÃØõ9c¹b®²?Rî€ÝÅ™'â„¿ha¬ì–<‡÷
r`2š"ÁÂ}wÄð¶;K~­2$ðhˆA<oÂÖ’<Âè¨²½ðf“jë ôF&“±]Š Æ|ÅrŠjFìº$ôùH¨%®	Zû7h“È¯ŠéP˜ b]ÎñWù¹·‰Y!<Ý†ÁÚÄ9ÚÞÏQÔ51Åôý‰=Mf7Öÿ}åF,œÞ°a(Uðö)n¸Ÿ£Òí¯™¬Y/ë·ª·Ö	ZBK)1nð1¸`mbáËÞ^vÿF÷X4Õ¢¼­Æs"AuL}ãìúl»ØX–u{†{ bChVŽ×ÎcÂ½¥ÄuCÊæÑ’Ì‡ùôãæç§mLešö')ÿ¨½§³üðÕ4švµ›]3ÖXˆ¨ŸF¼sï2îèƒXøÆÒío±µa2L<X¡lÑ!Qütn;‘î£ô6ÿk:ÂžÕ¦0þäC×b¯'U‚üOëO–¢7’v#’VQ›5ÔzÙëÇÙ’[‰ÝW’vÄü$ÜžŒOÙ»Løèt7v·,7DÃñSéifÂÚŽ™G’HýÀXTcèö›b‚M"ÀÒƒáË¢2• ùŒL±øDŒÕö„~nÉE¶äó-]p7S½Y¹ŒôöšÚ®õå’ÉXÖÑäŠ£_jlCáF¼úm
­P´ªCP°lƒÅôübÂëhj™'6,Š™uoBûãÙA“M©´ÜiÉZÄk““Ü¨ÓÉáËÌ…vóuã4Op†øÊ	Kr^‡W§·msWZ¹Ùq”KvñrI¡#éá¦ùÇ~¦á	¯ïÄÁHMýboKƒ¢ËÝ'ùtPŒZ+/Ppô•J~Ïìgºˆqc 
:n&ŒG¿-w<–Á¤‘>Õ”ø¤úg	‡iù-³(â ùÆý“nxu„G2&Azò£a“jPÌ„¤VˆóPí³Ad~[æ`ƒ!B+BÖY`Vùûþº*ênÓâõ¯#`¶³ÖàÖð_0‡Ò`­ñ
ûÔâj4OjS/Îàèßg%¼Ò	hÿqñfÞëoleÜoebñŠ}ñÆ&7O1Þ¸Ëù*ŒòÛ–Œ[-½ËÞ	èyŒp:%ü=¢™Í±¥ê‹.ŽQ`A3c¬dvu«ÝLQUeµ¢3äÆZÀ…«~¦qžY/aNÆkï³Àñ’fÜ„Û‚-P%ìTa;iÒ9LàÞÀ@³+4é%¤äãgQhËé»ºÀåŽ)jý³EÅt|ÚÒÂE…íðfD+ôR‹aÚ3Í±ñÄ³?Ù…i€h¿k:Þõm¬ð²~M[™®c¥-t‰Y¯]àà×7Ë=µ*ÇÀ¯)—ôÉö.Rô´•­³ÍÇ×.{ý]PLðŸ8¿-B¬Áþ8ö¨°þ'Žô$h°é(Â9Çà”¤ßÕ·ï«bF»àåT)óGöõwŸœb;{ÄžÊ wñSž±òXÿ_BôW‚³k•U"HR9}{àY%œlFæ£ðS?sÍÇ#ôRòtÅéÛ÷ÐXVoF¬’¦Ÿº˜ä6*4…	}?6ÂKp1ÁOÑŒœ|ôlþyxj1QGˆFN>:zt´s„—E6+([Ècõ‡vR^T^¿›²Ðn¦ðû'Š?«^½ì_å—ùˆáJB{OËAwˆÿž­e+Øh½²–ù»$›/dY·#_·?ÃœD¤Ð©¸tÖ²ŽÐ¯oƒ7Ù€ñ†w™Ýaû.TÉÐôþšáç¥\‚¬‚ÛPJó5ôÏ¾	ôîù>yf§þ-ÂµÓŠl¬)3]`ú&k’1'y	Û”eâœ2Uø¼±™ÿYÀ8~˜–—àÑ½½›ÑÂš¼?{…Ùáÿ7XþHeÒ/Ûn¡6L3‹Ô€ÂçAzé°?Ýþg:å‹Y8&‹óÔO¹u3P½¡ÉSÔs°@#†mv~W6¯Òìé¸ß?ÞÊMGY<²5Ð„é'Xè>Ch*¼ÿ†‰ã½Ò:¿½¤?ìïÑøù¾ÏCÓÇ˜+SmŸçy ª1lÿÐ*1ï2Î
JÎI¿¥ƒ·'Ï
WW
0ÛüMÁ2jßÂ¬€Ï`w=6-„zÜ\Ãp n«GõdT×ÅðÃ}•£(bÆëiu-À”]ï«Oá*DÕßá|?Séœ˜Áùüø¥HÝUÿúÌ!ö[{SUŸ	ÒdM•4£UÊM%Ü†ùµ¹ÈÅK#X~%VÉÙû0bƒÂiÉöëI}çtù6lk¶\ùI¡&&‰A¾h²H2ÃXiÝÀ4V§ÃØÝüé¿   ÿÿ B{Ì