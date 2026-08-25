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
  'CPSMS', 'CME', 'Cl√≠nica Geral', 'Higieniza√ß√£o', 'Dire√ß√£o', 
  'Recep√ß√£o', 'SAME', 'Copa', 'Administrativo', 'TI', 'Regula√ß√£o'
];

const SECTOR_COLORS: Record<string, string> = {
  'CPSMS': '#0284c7',
  'CME': '#7c3aed',
  'Cl√≠nica Geral': '#059669',
  'Higieniza√ß√£o': '#6366f1',
  'Dire√ß√£o': '#ef4444',
  'Recep√ß√£o': '#14b8a6',
  'SAME': '#7c2d12',
  'Copa': '#84cc16',
  'Administrativo': '#8b5cf6',
  'TI': '#1e293b',
  'Regula√ß√£o': '#fb923c'
};

const ROOMS = ['Sala A', 'Sala B', 'Almoxarifado Principal', 'Farm√°cia'];

const CATEGORY_COLORS: Record<string, string> = {
  'Odontol√≥gico': '#0284c7',
  'M√©dico Hospitalar': '#ef4444',
  'Aliment√≠cio': '#f59e0b',
  'Expediente': '#3b82f6',
  'Higiene': '#10b981',
  'Radiol√≥gico': '#8b5cf6',
  'Saneante': '#06b6d4',
  'Copa & Cozinha': '#f97316',
  'Papelaria': '#0ea5e9',
  'EPI': '#ec4899',
  'Gr√°fica': '#fbbf24',
  'Inform√°tica': '#6366f1',
  'Limpeza': '#059669',
  'Anest√©sico': '#7c3aed',
  'Medicamentos': '#be123c',
  'Fisioter√°picos': '#14b8a6',
  'Manuten√ß√£o': '#57534e',
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
                  A cota di√°ria do plano gratuito do Firebase (leituras de banco de dados) foi temporariamente atingida. 
                  O sistema continua armazenando suas altera√ß√µes e ativou a acelera√ß√£o por cache local no navegador.
                </p>
                <div className="bg-amber-50 border border-amber-200 p-4 rounded-2xl mb-6">
                  <p className="text-xs font-bold text-amber-800 mb-1">üí° O que voc√™ pode fazer:</p>
                  <ul className="text-[11px] text-amber-700 space-y-1 list-disc list-inside">
                    <li>Recarregue a p√°gina para utilizar os dados em cache no seu navegador.</li>
                    <li>As cotas di√°rias gratuitas s√£o renovadas automaticamente pelo Google Firebase a cada novo ciclo di√°rio.</li>
                  </ul>
                </div>
              </>
            ) : (
              <>
                <h2 className="text-2xl font-black text-center text-slate-900 mb-4">Algo deu errado</h2>
                <p className="text-slate-500 text-center mb-6 text-sm">
                  Ocorreu um erro inesperado. Por favor, recarregue a p√°gina ou tente novamente.
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
              Recarregar Aplica√ß√£o
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
  const [quantitativoCategory, setQuantitativoCategory] = useState('Material M√©dico-Hospitalar');
  const [quantitativoTitle, setQuantitativoTitle] = useState('');
  const [quantitativoCriticalAnalysis, setQuantitativoCriticalAnalysis] = useState('');
  const [isEditingQuantitativoAnalysis, setIsEditingQuantitativoAnalysis] = useState(false);
  const quantitativoReportRef = useRef<HTMLDivElement>(null);
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'inventory' | 'history' | 'requests' | 'admin-devolutions' | 'reports' | 'my-requests' | 'new-request' | 'devolution' | 'users' | 'trash' | 'leader-stats'>('dashboard');
  const leaderStatistics = useMemo(() => {
    if (userProfile?.role !== 'L√çDER' && userProfile?.role !== 'SETOR') return { topRequested: [], topDelivered: [] };

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
  const [devolutionReason, setDevolutionReason] = useState('N√£o teve uso');
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
  const [inventoryLocation, setInventoryLocation] = useState<'Almoxarifado' | 'Farm√°cia'>('Almoxarifado');

  const isAdmin = userProfile?.role === 'ADMIN' || 
                  user?.email === 'gerlianemagalhaes79@gmail.com' || 
                  userProfile?.sector === 'Almoxarifado';

  useEffect(() => {
    if (userProfile?.sector === 'Farm√°cia' || selectedSector === 'Farm√°cia') {
      if (!isAdmin) {
        setInventoryLocation('Farm√°cia');
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
  const [requestBasket, setRequestBasket] = useState<{product_id: string, product_name: string, quantity: number, batch_id?: string}[]>([]);
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
          console.log(`Auto-otimizando estoque m√≠nimo para ${updates.length} lotes...`);
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
          console.log("Otimiza√ß√£o de estoque m√≠nimo conclu√≠da.");
        } catch (error) {
          console.error("Erro ao auto-atualizar estoques m√≠nimos:", error);
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
  const [categories, setCategories] = useState<string[]>(['Odontol√≥gico', 'M√©dico Hospitalar', 'Aliment√≠cio', 'Expediente', 'Higiene', 'Radiol√≥gico', 'Saneante', 'Copa & Cozinha', 'Papelaria', 'EPI', 'Gr√°fica', 'Inform√°tica', 'Limpeza', 'Anest√©sico', 'Medicamentos', 'Fisioter√°picos', 'Manuten√ß√£o']);
  const [showNewCategoryInput, setShowNewCategoryInput] = useState(false);
  const [newCategoryName, setNewCategoryName] = useState('');

  useEffect(() => {
    if (showAddModal) {
      setBulkEntry({
        supplier: '',
        category: inventoryLocation === 'Farm√°cia' ? 'Medicamentos' : 'Expediente',
        origin: 'extra' as 'contract' | 'extra' | 'donation',
        room: inventoryLocation === 'Farm√°cia' ? 'Farm√°cia' : 'Almoxarifado Principal',
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
      showToast(`Pre√ßo unit√°rio de "${itemToUpdate.name}" atualizado em todos os lotes!`, "success");
      setEditingPrice(null);
    } catch (error: any) {
      handleFirestoreError(error, OperationType.UPDATE, `items/${editingPrice.id}`);
      showToast(`Erro ao atualizar pre√ßo: ${error.message}`, "error");
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
      showToast("Por favor, selecione ou informe uma categoria v√°lida.", "error");
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
      showToast("Digite CONFIRMAR, TESTE ou EXCLUIR para autorizar a remo√ß√£o.", "error");
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

        showToast(`${entryTrans.length} entradas de materiais de teste foram exclu√≠das com sucesso!`, "success");
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

        showToast(`Todas as entradas e itens de estoque cadastrados como teste foram exclu√≠dos!`, "success");
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

        showToast(`Todos os dados de teste (entradas, sa√≠das, requisi√ß√µes e estoque) foram completamente limpos!`, "success");
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
            showToast("Erro: E-mail n√£o encontrado no login do Google.", "error");
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
            showToast("Acesso negado: Seu e-mail n√£o est√° cadastrado no sistema. Entre em contato com o administrador.", "error");
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
          showToast(`Erro na autentica√ß√£o: ${error.message}`, "error");
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
        showToast("Erro: Dom√≠nio n√£o autorizado no Firebase Auth.", "error");
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
      
      showToast("Usu√°rio pr√©-cadastrado com sucesso! Agora ele pode entrar usando o Google.", "success");
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
          throw new Error("Movimenta√ß√£o n√£o encontrada.");
        }
        
        const transData = transSnap.data() as Transaction;

        if (transData.deletedAt) {
          throw new Error("Esta movimenta√ß√£o j√° foi exclu√≠da.");
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
      alert(`Erro ao excluir movimenta√ß√£o: ${error.message}`);
    }
  };

  const handleRecoverTransaction = async (id: string) => {
    if (!id) return;
    try {
      await runTransaction(db, async (transaction) => {
        const transRef = doc(db, 'transactions', id);
        const transSnap = await transaction.get(transRef);
        
        if (!transSnap.exists()) {
          throw new Error("Movimenta√ß√£o n√£o encontrada.");
        }
        
        const transData = transSnap.data() as Transaction;

        if (!transData.deletedAt) {
          throw new Error("Esta movimenta√ß√£o n√£o est√° exclu√≠da.");
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
      alert(`Erro ao recuperar movimenta√ß√£o: ${error.message}`);
    }
  };

  const handleRecoverAllTransactions = async () => {
    const deletedTrans = transactions.filter(t => !!t.deletedAt);
    if (deletedTrans.length === 0) return;
    
    if (!confirm(`Deseja restaurar todas as ${deletedTrans.length} movimenta√ß√µes exclu√≠das?`)) return;

    try {
      // We'll process them one by one to ensure stock is updated correctly via transactions
      for (const t of deletedTrans) {
        await handleRecoverTransaction(t.id);
      }
      alert("Todas as movimenta√ß√µes foram restauradas com sucesso!");
    } catch (error: any) {
      console.error("Error recovering all transactions:", error);
      alert(`Erro ao restaurar movimenta√ß√µes: ${error.message}`);
    }
  };

  const handleSubmitRequest = async () => {
    if (requestBasket.length === 0) {
      showToast("Adicione pelo menos um item √† solicita√ß√£o.", "error");
      return;
    }

    setIsSubmittingRequest(true);
    const loadingToast = showToast("Processando sua solicita√ß√£o...", "info");
    
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
            `Estoque insuficiente para "${originalName}". Dispon√≠vel: ${totalAvailable}.`, 
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
        requestData.requesterName = user?.displayName || user?.email || 'Usu√°rio';
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
          quantity_approved: Math.max(1, Math.floor(Number(item.quantity) || 1)),
          batch_id: item.batch_id || ''
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
                createNotification(d.id, 'Nova Solicita√ß√£o', `Setor ${selectedSector} enviou uma nova solicita√ß√£o.`, requestId);
                notified.add(d.id);
              }
            });
          };
          notify(adminSnap);
          notify(almoxSnap);
        } catch (notifErr) {
          console.warn("Falha ao enviar notifica√ß√µes:", notifErr);
        }
      }

      showToast(editingRequest ? "Altera√ß√µes salvas com sucesso!" : "Solicita√ß√£o enviada com sucesso!", "success");
      setRequestBasket([]);
      setRequestObservation('');
      setEditingRequest(null);
      setActiveTab('my-requests');
    } catch (error: any) {
      console.error("Erro cr√≠tico ao salvar:", error);
      showToast(`N√£o foi poss√≠vel salvar: ${error.message}. Verifique sua conex√£o e tente novamente.`, "error");
    } finally {
      setIsSubmittingRequest(false);
    }
  };

  const handleEditRequest = (request: MaterialRequest) => {
    setSelectedSector(request.sector);
    const reqItems = allRequestItems.filter(ri => ri.request_id === request.id);
    setRequestBasket(reqItems.map(i => ({
      product_id: i.product_id,
      product_name: i.product_name,
      quantity: i.quantity_requested,
      batch_id: i.batch_id || ''
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
    if (!window.confirm("Tem certeza que deseja enviar este item para a lixeira? Ele ser√° exclu√≠do definitivamente ap√≥s 3 dias.")) return;
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
      showToast("Nenhuma solicita√ß√£o pendente ou em separa√ß√£o encontrada para este per√≠odo.", "info");
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
        showToast("Status das solicita√ß√µes atualizado para 'Em Separa√ß√£o'!", "success");
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
          <title>Impress√£o de Solicita√ß√µes - ${periodStr}</title>
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
            const reqItems = allRequestItems.filter(ri => ri.request_id === req.id);
            return `
              <div class="request-card">
                <h1>Solicita√ß√£o de Material</h1>
                <table class="header-table">
                  <tr>
                    <td><strong>N√∫mero:</strong> #${req.id.slice(-5).toUpperCase()}</td>
                    <td><strong>Data de Cria√ß√£o:</strong> ${new Date(req.date).toLocaleDateString('pt-BR')}</td>
                  </tr>
                  <tr>
                    <td><strong>Setor Solicitante:</strong> ${req.sector}</td>
                    <td><strong>Status:</strong> EM SEPARA√á√ÉO</td>
                  </tr>
                  <tr>
                    <td colspan="2"><strong>Solicitante:</strong> ${req.requesterEmail}</td>
                  </tr>
                  ${req.observation ? `<tr><td colspan="2"><strong>Observa√ß√µes:</strong> ${req.observation}</td></tr>` : ''}
                </table>

                <h3 style="margin: 6px 0 3px 0; font-size: 9px; border-bottom: 1px solid #1C1917; padding-bottom: 2px; text-transform: uppercase;">ITENS DA SOLICITA√á√ÉO (Para separa√ß√£o f√≠sica)</h3>
                <table class="items-table">
                  <thead>
                    <tr>
                      <th>Produto / Descri√ß√£o</th>
                      <th style="width: 70px; text-align: center;">Qtd Solicitada</th>
                      <th class="blank-col">Qtd Separada</th>
                      <th>Obs. / Lote do Material</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${reqItems.map(item => {
                      const activeBatches = items.filter((it: any) => !it.deletedAt && normalizeString(it.name) === normalizeString(item.product_name) && (it.quantity || 0) > 0);
                      const hasMultipleBatches = activeBatches.length > 1;
                      const selectedBatch = item.batch_id ? items.find((it: any) => it.id === item.batch_id) : null;
                      
                      let lotesInfoHtml = '';
                      if (selectedBatch) {
                        lotesInfoHtml = `<div style="font-size: 8px; color: #1D4ED8; font-weight: bold; margin-top: 2px;">‚Ä¢ Lote Solicitado: <strong>${selectedBatch.batch_number || 'Sem Lote'}</strong> (Disp: ${selectedBatch.quantity} un | Venc: ${selectedBatch.expiry_date || 'Indet.'})</div>`;
                      } else if (hasMultipleBatches) {
                        lotesInfoHtml = `
                          <div style="font-size: 8px; color: #B45309; font-weight: bold; margin-top: 2px; background-color: #FEF3C7; padding: 2px 4px; border-radius: 3px; display: inline-block;">
                            ‚ö†Ô∏è Aten√ß√£o: ${activeBatches.length} Lotes no Estoque
                          </div>
                          <div style="font-size: 7.5px; color: #44403C; margin-top: 2px; padding-left: 4px; border-left: 2px solid #F59E0B;">
                            ${activeBatches.map((b: any) => `‚Ä¢ Lote: <strong>${b.batch_number || 'S/L'}</strong> (Disp: ${b.quantity} un | Venc: ${b.expiry_date || 'Indet.'})`).join('<br/>')}
                          </div>
                        `;
                      } else if (activeBatches.length === 1 && activeBatches[0].batch_number) {
                        lotesInfoHtml = `<div style="font-size: 8px; color: #57534E; margin-top: 2px;">‚Ä¢ Lote: ${activeBatches[0].batch_number} (Disp: ${activeBatches[0].quantity} un | Venc: ${activeBatches[0].expiry_date || 'Indet.'})</div>`;
                      }

                      let obsLoteCol = '';
                      if (selectedBatch) {
                        obsLoteCol = `<strong>${selectedBatch.batch_number || '---'}</strong>`;
                      } else if (hasMultipleBatches) {
                        obsLoteCol = `<span style="font-size: 7.5px; color: #B45309; font-weight: bold;">[M√∫ltiplos Lotes]<br/>Anotar Lote: _________</span>`;
                      } else if (activeBatches.length === 1) {
                        obsLoteCol = activeBatches[0].batch_number || '---';
                      }

                      return `
                        <tr>
                          <td style="font-size: 8.5px; vertical-align: top;">
                            <strong>${item.product_name}</strong>
                            ${lotesInfoHtml}
                          </td>
                          <td style="text-align: center; font-size: 8.5px; font-weight: bold; vertical-align: top;">${item.quantity_requested}</td>
                          <td class="blank-col" style="border-bottom: 1px solid #1C1917; vertical-align: top;"></td>
                          <td style="border-bottom: 1px solid #1C1917; font-size: 8px; vertical-align: top;">${obsLoteCol}</td>
                        </tr>
                      `;
                    }).join('')}
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
      showToast("N√£o √© poss√≠vel excluir uma solicita√ß√£o que j√° foi entregue.", "error");
      return;
    }
    
    if (!window.confirm("Tem certeza que deseja enviar esta solicita√ß√£o para a lixeira? Ela ser√° exclu√≠da definitivamente ap√≥s 3 dias.")) return;
    try {
      await updateDoc(doc(db, 'requests', requestId), {
        deletedAt: new Date().toISOString(),
        deletedBy: user?.email
      });
      showToast("Solicita√ß√£o enviada para a lixeira.", "success");
    } catch (error: any) {
      handleFirestoreError(error, OperationType.UPDATE, `requests/${requestId}`);
      showToast(`Erro ao excluir solicita√ß√£o: ${error.message}`, "error");
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
      showToast(`"${productName}" adicionado √† solicita√ß√£o.`, "success");
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
      showToast("Observa√ß√£o atualizada com sucesso!", "success");
    } catch (error: any) {
      console.error("Error updating observation:", error);
      showToast(`Erro ao atualizar observa√ß√£o: ${error.message}`, "error");
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
        showToast("Status alterado para 'Em Separa√ß√£o'!", "success");
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

    const reqItems = allRequestItems.filter(ri => ri.request_id === request.id);
    const dateStr = new Date(request.date).toLocaleDateString('pt-BR');

    const content = `
      <html>
        <head>
          <title>Solicita√ß√£o de Material - #${request.id.slice(-5).toUpperCase()}</title>
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
          <h1>Solicita√ß√£o de Material</h1>
          <table class="header-table">
            <tr>
              <td><strong>N√∫mero:</strong> #${request.id.slice(-5).toUpperCase()}</td>
              <td><strong>Data:</strong> ${dateStr}</td>
            </tr>
            <tr>
              <td><strong>Setor Solicitante:</strong> ${request.sector}</td>
              <td><strong>Status:</strong> ${request.status === 'PENDENTE' ? 'PENDENTE' : 'EM SEPARA√á√ÉO'}</td>
            </tr>
            <tr>
              <td colspan="2"><strong>Solicitante:</strong> ${request.requesterEmail}</td>
            </tr>
            ${request.observation ? `<tr><td colspan="2"><strong>Observa√ß√µes do Solicitante:</strong> ${request.observation}</td></tr>` : ''}
          </table>

          <h3 style="margin-top: 30px; font-size: 16px; border-bottom: 1px solid #1C1917; padding-bottom: 5px;">ITENS DA SOLICITA√á√ÉO (Para separa√ß√£o f√≠sica)</h3>
          <table class="items-table">
            <thead>
              <tr>
                <th>Produto / Descri√ß√£o</th>
                <th style="width: 100px; text-align: center;">Qtd Solicitada</th>
                <th class="blank-col">Qtd Separada (Anotar)</th>
                <th>Obs. / Lote do Material</th>
              </tr>
            </thead>
            <tbody>
              ${reqItems.map(item => {
                const activeBatches = items.filter((it: any) => !it.deletedAt && normalizeString(it.name) === normalizeString(item.product_name) && (it.quantity || 0) > 0);
                const hasMultipleBatches = activeBatches.length > 1;
                const selectedBatch = item.batch_id ? items.find((it: any) => it.id === item.batch_id) : null;
                
                let lotesInfoHtml = '';
                if (selectedBatch) {
                  lotesInfoHtml = `<div style="font-size: 11px; color: #1D4ED8; font-weight: bold; margin-top: 4px;">‚Ä¢ Lote Solicitado: <strong>${selectedBatch.batch_number || 'Sem Lote'}</strong> (Disp: ${selectedBatch.quantity} un | Venc: ${selectedBatch.expiry_date || 'Indeterminado'})</div>`;
                } else if (hasMultipleBatches) {
                  lotesInfoHtml = `
                    <div style="font-size: 11px; color: #B45309; font-weight: bold; margin-top: 4px; background-color: #FEF3C7; padding: 4px 8px; border-radius: 4px; display: inline-block;">
                      ‚ö†Ô∏è Aten√ß√£o: Material com ${activeBatches.length} Lotes no Estoque
                    </div>
                    <div style="font-size: 10.5px; color: #44403C; margin-top: 4px; padding-left: 6px; border-left: 3px solid #F59E0B;">
                      ${activeBatches.map((b: any) => `‚Ä¢ Lote: <strong>${b.batch_number || 'S/L'}</strong> (Saldo: ${b.quantity} un | Venc: ${b.expiry_date || 'Indeterminado'})`).join('<br/>')}
                    </div>
                  `;
                } else if (activeBatches.length === 1 && activeBatches[0].batch_number) {
                  lotesInfoHtml = `<div style="font-size: 11px; color: #57534E; margin-top: 3px;">‚Ä¢ Lote Dispon√≠vel: ${activeBatches[0].batch_number} (Saldo: ${activeBatches[0].quantity} un | Venc: ${activeBatches[0].expiry_date || 'Indet.'})</div>`;
                }

                let obsLoteCol = '';
                if (selectedBatch) {
                  obsLoteCol = `<strong>${selectedBatch.batch_number || '---'}</strong>`;
                } else if (hasMultipleBatches) {
                  obsLoteCol = `<span style="font-size: 10px; color: #B45309; font-weight: bold;">[M√∫ltiplos Lotes]<br/>Anotar Lote: _____________</span>`;
                } else if (activeBatches.length === 1) {
                  obsLoteCol = activeBatches[0].batch_number || '---';
                }

                return `
                  <tr>
                    <td style="font-size: 12px; vertical-align: top;">
                      <strong>${item.product_name}</strong>
                      ${lotesInfoHtml}
                    </td>
                    <td style="text-align: center; font-size: 13px; font-weight: bold; vertical-align: top;">${item.quantity_requested}</td>
                    <td class="blank-col" style="border-bottom: 1px solid #1C1917; vertical-align: top;"></td>
                    <td style="border-bottom: 1px solid #1C1917; font-size: 11px; vertical-align: top;">${obsLoteCol}</td>
                  </tr>
                `;
              }).join('')}
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
              Respons√°vel pela Separa√ß√£o (Almoxarifado)
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
      showToast("Processando aprova√ß√£o e baixa no estoque...", "info");
      
      const requestRef = doc(db, 'requests', requestId);
      const requestSnap = await getDoc(requestRef);
      if (!requestSnap.exists()) throw new Error("Solicita√ß√£o n√£o encontrada.");
      const requestData = requestSnap.data() as MaterialRequest;

      if (requestData.status === 'ENTREGUE') {
        showToast("Esta solicita√ß√£o j√° foi entregue.", "info");
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
        if (reqItem.batch_id) {
          const chosenBatch = batches.find(b => b.id === reqItem.batch_id);
          if (chosenBatch) {
            batches = [chosenBatch, ...batches.filter(b => b.id !== reqItem.batch_id)];
          }
        }

        let pharmItems: any[] = [];
        if (requestData.sector === 'Farm√°cia') {
          pharmItems = allActiveItems
            .filter(item => normalizeString(item.name) === normalizedReqName && item.location === 'Farm√°cia')
            .map(item => ({ id: item.id, batch_number: item.batch_number, ref: doc(db, 'items', item.id) }));
        }

        itemsStockData.push({ reqItem, batches, pharmItems });
      }

      const itemBatchesAllocated: Record<string, string[]> = {};

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

            if (batch.batch_number) {
              if (!itemBatchesAllocated[reqItem.product_name]) {
                itemBatchesAllocated[reqItem.product_name] = [];
              }
              if (!itemBatchesAllocated[reqItem.product_name].includes(batch.batch_number)) {
                itemBatchesAllocated[reqItem.product_name].push(batch.batch_number);
              }
            }

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

            if (requestData.sector === 'Farm√°cia' && batch.location !== 'Farm√°cia') {
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
                  supplier: batch.supplier || 'Transfer√™ncia',
                  batch_number: batch.batch_number || '',
                  expiry_date: batch.expiry_date || 'Indeterminada',
                  initial_quantity: toTake,
                  quantity: toTake,
                  min_quantity: batch.min_quantity || 0,
                  unit_price: batch.unit_price || 0,
                  location: 'Farm√°cia',
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
      showToast("Solicita√ß√£o aprovada, entregue e estoque baixado automaticamente!", "success");
      setShowRequestDetailModal({ show: false });

      // Notifications
      const uSnap = await getDocs(query(collection(db, 'users'), where('email', '==', requestData.requesterEmail)));
      if (!uSnap.empty) {
        await createNotification(uSnap.docs[0].id, 'Solicita√ß√£o Entregue', `Sua solicita√ß√£o #${requestId.slice(-5).toUpperCase()} foi aprovada e entregue.`, requestId);
      }

      // Stock Zero Notifications
      for (const { reqItem } of itemsStockData) {
        await checkStockAndNotify(reqItem.product_name);
      }

      // Receipt
      const itemsForReceipt = currentRequestItems.filter(i => i.quantity_approved > 0).map(i => ({
        product_name: i.product_name,
        quantity: i.quantity_approved,
        batch_number: (itemBatchesAllocated[i.product_name] && itemBatchesAllocated[i.product_name].length > 0)
          ? itemBatchesAllocated[i.product_name].join(', ')
          : (items.find(it => it.name === i.product_name && it.batch_number)?.batch_number || '---')
      }));
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
      showToast("Por favor, adicione pelo menos um item √† devolu√ß√£o.", "info");
      return;
    }

    // Validate quantities
    for (const item of devolutionBasket) {
      if (item.quantity <= 0) {
        showToast(`Por favor, insira uma quantidade maior que zero para ${item.product_name}.`, "error");
        return;
      }
      if (item.quantity > item.maxQty) {
        showToast(`Quantidade inv√°lida para ${item.product_name}. M√°ximo permitido: ${item.maxQty}`, "error");
        return;
      }
    }

    try {
      setIsProcessingDevolution(true);
      showToast("Enviando solicita√ß√£o de devolu√ß√£o...", "info");

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
        requesterName: userProfile?.name || user?.displayName || user?.email || 'Usu√°rio',
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
            title: 'Solicita√ß√£o de Devolu√ß√£o',
            message: `Setor ${selectedSector} solicitou devolu√ß√£o de materiais.`,
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

      showToast("Solicita√ß√£o de devolu√ß√£o enviada para o almoxarifado!", "success");
      setShowDevolutionModal({ show: false });
      setDevolutionBasket([]);
      setDevolutionObservation('');
      
      if (showRequestDetailModal.show && showDevolutionModal.request && showRequestDetailModal.request?.id === showDevolutionModal.request.id) {
        setShowRequestDetailModal({ show: false });
      }

    } catch (error: any) {
      console.error("Erro ao solicitar devolu√ß√£o:", error);
      showToast(`Erro ao solicitar devolu√ß√£o: ${error.message}`, "error");
    } finally {
      setIsProcessingDevolution(false);
    }
  };

  const handleApproveDevolution = async (requestId: string, devItems: RequestItem[]) => {
    try {
      setIsProcessingDevolution(true);
      showToast("Aprovando devolu√ß√£o e retornando ao estoque...", "info");

      // Fetch active stock items
      const itemsSnapshot = await getDocs(collection(db, 'items'));
      const allActiveItems = itemsSnapshot.docs
        .map(d => ({ id: d.id, ...d.data() } as Item))
        .filter(i => !i.deletedAt);

      await runTransaction(db, async (transaction) => {
        const requestRef = doc(db, 'requests', requestId);
        const requestSnap = await transaction.get(requestRef);
        if (!requestSnap.exists()) throw new Error("Solicita√ß√£o de devolu√ß√£o n√£o encontrada.");
        const requestData = requestSnap.data() as MaterialRequest;

        if (requestData.status === 'DEVOLUCAO_APROVADA') {
          throw new Error("Esta devolu√ß√£o j√° foi aprovada anteriormente.");
        }

        // Collect all doc IDs that we need to read in the transaction:
        // 1) Sector source items (Farm√°cia / Requesting sector)
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

          // 1. DECREASE stock in sector (e.g., Farm√°cia)
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
          let batchNumber = sourceItemDoc?.data.batch_number || 'Devolu√ß√£o';
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
            returnReason: requestData.returnReason || 'N√£o especificado',
            observation: requestData.observation || ''
          });

          // 4. Log exit transaction for Sector/Farm√°cia
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
            expiryReason: requestData.returnReason || 'Devolu√ß√£o ao Almoxarifado',
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
            title: 'Devolu√ß√£o Aprovada',
            message: `Sua solicita√ß√£o de devolu√ß√£o para o setor ${requestData.sector} foi aprovada. Os materiais retornaram ao estoque.`,
            date: new Date().toISOString(),
            read: false,
            requestId: requestId,
            type: 'REQUEST'
          });
        }
      }

      showToast("Devolu√ß√£o aprovada com sucesso! Materiais retornados ao estoque.", "success");
      setShowRequestDetailModal({ show: false });

    } catch (error: any) {
      console.error("Erro ao aprovar devolu√ß√£o:", error);
      showToast(`Erro ao aprovar devolu√ß√£o: ${error.message}`, "error");
    } finally {
      setIsProcessingDevolution(false);
    }
  };

  const handleRejectDevolution = async (requestId: string) => {
    try {
      setIsProcessingDevolution(true);
      showToast("Recusando devolu√ß√£o...", "info");

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
            title: 'Devolu√ß√£o Recusada',
            message: `Sua solicita√ß√£o de devolu√ß√£o para o setor ${requestData.sector} foi recusada pelo almoxarifado.`,
            date: new Date().toISOString(),
            read: false,
            requestId: requestId,
            type: 'REQUEST'
          });
        }
      }

      showToast("Devolu√ß√£o recusada com sucesso.", "success");
      setShowRequestDetailModal({ show: false });
    } catch (error: any) {
      console.error("Erro ao recusar devolu√ß√£o:", error);
      showToast(`Erro ao recusar devolu√ß√£o: ${error.message}`, "error");
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
            ? `Sua solicita√ß√£o #${requestId.slice(-5).toUpperCase()} foi aprovada. Obs: ${adminObservation}`
            : `Sua solicita√ß√£o #${requestId.slice(-5).toUpperCase()} foi aprovada.`;
          await createNotification(userSnap.docs[0].id, 'Solicita√ß√£o Aprovada', msg, requestId);
        }
      }

      showToast("Solicita√ß√£o aprovada!", "success");
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
      if (!requestSnap.exists()) throw new Error("Solicita√ß√£o n√£o encontrada.");
      const requestData = requestSnap.data() as MaterialRequest;

      if (requestData.status === 'ENTREGUE') {
        showToast("Esta solicita√ß√£o j√° foi entregue.", "info");
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
        if (reqItem.batch_id) {
          const chosenBatch = batches.find(b => b.id === reqItem.batch_id);
          if (chosenBatch) {
            batches = [chosenBatch, ...batches.filter(b => b.id !== reqItem.batch_id)];
          }
        }

        let pharmItems: any[] = [];
        if (requestData.sector === 'Farm√°cia') {
          pharmItems = allActiveItems
            .filter(item => normalizeString(item.name) === normalizedReqName && item.location === 'Farm√°cia')
            .map(item => ({ id: item.id, batch_number: item.batch_number, ref: doc(db, 'items', item.id) }));
        }

        itemsStockData.push({ reqItem, batches, pharmItems });
      }

      const itemBatchesAllocated: Record<string, string[]> = {};

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

            if (batch.batch_number) {
              if (!itemBatchesAllocated[reqItem.product_name]) {
                itemBatchesAllocated[reqItem.product_name] = [];
              }
              if (!itemBatchesAllocated[reqItem.product_name].includes(batch.batch_number)) {
                itemBatchesAllocated[reqItem.product_name].push(batch.batch_number);
              }
            }

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

            if (requestData.sector === 'Farm√°cia' && batch.location !== 'Farm√°cia') {
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
                  supplier: batch.supplier || 'Transfer√™ncia',
                  batch_number: batch.batch_number || '',
                  expiry_date: batch.expiry_date || 'Indeterminada',
                  initial_quantity: toTake,
                  quantity: toTake,
                  min_quantity: batch.min_quantity || 0,
                  unit_price: batch.unit_price || 0,
                  location: 'Farm√°cia',
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
        await createNotification(uSnap.docs[0].id, 'Entrega Realizada', `Sua solicita√ß√£o #${requestId.slice(-5).toUpperCase()} foi entregue.`, requestId);
      }

      // Stock Zero Notifications
      for (const { reqItem } of itemsStockData) {
        await checkStockAndNotify(reqItem.product_name);
      }

      // Receipt
      const itemsForReceipt = requestItems.filter(i => i.quantity_approved > 0).map(i => ({
        product_name: i.product_name,
        quantity: i.quantity_approved,
        batch_number: (itemBatchesAllocated[i.product_name] && itemBatchesAllocated[i.product_name].length > 0)
          ? itemBatchesAllocated[i.product_name].join(', ')
          : (items.find(it => it.name === i.product_name && it.batch_number)?.batch_number || '---')
      }));
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
            ? `Sua solicita√ß√£o #${requestId.slice(-5).toUpperCase()} foi recusada. Motivo: ${adminObservation}`
            : `Sua solicita√ß√£o #${requestId.slice(-5).toUpperCase()} foi recusada.`;
          await createNotification(userSnap.docs[0].id, 'Solicita√ß√£o Recusada', msg, requestId);
        }
      }
      
      showToast("Solicita√ß√£o recusada.", "success");
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
          showToast("O nome do produto n√£o pode estar vazio ou conter apenas espa√ßos.", "error");
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
              throw new Error("Item n√£o encontrado durante a atualiza√ß√£o.");
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
              throw new Error(`Item ${b.item_id} n√£o encontrado.`);
            }

            const currentItemData = itemSnap.data() as Item;
            const currentQty = Number(currentItemData.quantity) || 0;
            if (currentQty < b.quantity) {
              throw new Error(`Estoque insuficiente para o item ${currentItemData.name}. Dispon√≠vel: ${currentQty}`);
            }

            let pharmacyItemSnap = null;
            if (selectedSector === 'Farm√°cia' && exitReason === 'consumo') {
              const pharmacyItemsQuery = query(
                collection(db, 'items'),
                where('name', '==', currentItemData.name),
                where('batch_number', '==', currentItemData.batch_number || ''),
                where('location', '==', 'Farm√°cia')
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

            const sectorValue = modalSector || (inventoryLocation === 'Farm√°cia' ? 'Farm√°cia (Consumo Interno)' : 'Almoxarifado');

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
              donationUnitName: exitReason === 'doacao' ? (donationUnitName || 'CEO - Centro de Especialidades Odontol√≥gicas') : null,
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
                  location: 'Farm√°cia',
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
                location: 'Farm√°cia',
                date: new Date().toISOString(),
                responsible: 'Sistema (Transfer√™ncia)',
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
            throw new Error("Item n√£o encontrado.");
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
            product_name: foundItem?.name || 'Produto N√£o Identificado',
            quantity: b.quantity,
            batch_number: foundItem?.batch_number || '',
            expiry_date: foundItem?.expiry_date || ''
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
            donatingUnitName: donationUnitName || 'CEO - Centro de Especialidades Odontol√≥gicas',
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
      console.error('Erro na transa√ß√£o:', error);
      alert(`Erro na movimenta√ß√£o: ${error.message}`);
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
            'Item': `   ‚Ü≥ ${item.name}`,
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
      XLSX.utils.book_append_sheet(wb, ws, "Relat√≥rio de Sa√≠das");

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
          if (group.durationWeeks <= 4) status = 'MUITO CR√çTICO';
          else if (group.durationWeeks <= 8) status = 'CR√çTICO';
        }
        
        return {
          'Item': group.name,
          'Categoria': group.category || '---',
          'Estoque Total': group.total_quantity,
          'M√≠nimo': group.min_quantity,
          'Dura√ß√£o (Semanas)': group.durationWeeks === 'infinite' ? '‚àû' : group.durationWeeks.toFixed(1),
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
        'Relat√≥rio de Estoque Atual',
        `Gerado em: ${format(new Date(), 'dd/MM/yyyy HH:mm')}`
      );

      // Prepare data for table
      const tableData = groupedArray.map(group => {
        let status = group.total_quantity <= group.min_quantity ? 'BAIXO' : 'OK';
        if (group.durationWeeks !== 'infinite') {
          if (group.durationWeeks <= 4) status = 'MUITO CR√çTICO';
          else if (group.durationWeeks <= 8) status = 'CR√çTICO';
        }
        
        return [
          group.name,
          group.category || '---',
          group.total_quantity.toString(),
          group.durationWeeks === 'infinite' ? '‚àû' : group.durationWeeks.toFixed(1),
          group.min_quantity.toString(),
          status
        ];
      });
      
      // Generate table
      autoTable(doc, {
        startY: startY + 4,
        head: [['Item', 'Categoria', 'Estoque', 'Dura√ß√£o (Sem)', 'M√≠nimo', 'Status']],
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
            if (text === 'BAIXO' || text === 'MUITO CR√çTICO') {
              data.cell.styles.textColor = [225, 29, 72]; // rose-600
              data.cell.styles.fontStyle = 'bold';
            } else if (text === 'CR√çTICO') {
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
      const locationLabel = inventoryLocation === 'Farm√°cia' ? 'Farm√°cia (Medicamentos)' : 'Almoxarifado Geral';
      
      const startY = drawPDFLetterhead(
        doc,
        `RELAT√ìRIO DE ITENS CR√çTICOS ‚Äî ESTOQUE BAIXO`,
        `Unidade: ${locationLabel} ‚Ä¢ Data de Emiss√£o: ${dateStr}`
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
        doc.text(`Todos os insumos cadastrados no ${locationLabel} est√£o acima do estoque m√≠nimo.`, 14, startY + 20);

        const dateFileStr = format(new Date(), 'dd-MM-yyyy');
        doc.save(`Relatorio_Estoque_Baixo_${inventoryLocation}_${dateFileStr}.pdf`);
        showToast("Relat√≥rio gerado: Nenhum item com estoque baixo encontrado.", "info");
        return;
      }

      // Prepare table data
      const tableData = lowStockGroupsList.map(group => {
        const deficit = Math.max(0, group.min_quantity - group.total_quantity);
        let status = 'ESTOQUE BAIXO';
        if (group.total_quantity === 0) {
          status = 'ZERADO / SEM ESTOQUE';
        } else if (group.total_quantity <= (group.min_quantity * 0.5)) {
          status = 'MUITO CR√çTICO';
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
        head: [['Material / Medicamento', 'Categoria', 'Estoque Atual', 'Estoque M√≠nimo', 'D√©ficit (Reposi√ß√£o)', 'Status Cr√≠tico']],
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
            if (text.includes('ZERADO') || text === 'MUITO CR√çTICO') {
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
      doc.text(`Total de itens identificados com estoque baixo ou cr√≠tico: ${lowStockGroupsList.length}`, 14, finalY);

      const dateFileStr = format(new Date(), 'dd-MM-yyyy');
      doc.save(`Relatorio_Estoque_Baixo_${inventoryLocation}_${dateFileStr}.pdf`);
      showToast("Relat√≥rio PDF de estoque baixo gerado com sucesso!", "success");
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
        'Cat√°logo de Materiais em Estoque',
        `CEO - Centro de Especialidades Odontol√≥gicas ‚Ä¢ Gerado em: ${format(new Date(), 'dd/MM/yyyy HH:mm')}`
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
          `P√°gina ${i} de ${pageCount} - Cat√°logo gerado para consulta administrativa`,
          pageWidth / 2,
          doc.internal.pageSize.height - 10,
          { align: 'center' }
        );
      }
      
      const dateStr = format(new Date(), 'dd-MM-yyyy');
      doc.save(`Catalogo_Materiais_${dateStr}.pdf`);
      showToast("Cat√°logo de materiais exportado com sucesso!", "success");
    } catch (error) {
      console.error('Erro ao exportar cat√°logo:', error);
      showToast("Erro ao exportar cat√°logo de materiais.", "error");
    }
  };

  const handleExportRequestsPDF = () => {
    try {
      const doc = new jsPDF();
      
      const startY = drawPDFLetterhead(
        doc,
        'Relat√≥rio de Solicita√ß√µes de Materiais',
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
        head: [['N¬∫', 'Data', 'Setor', 'Status', 'Itens']],
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
        setModalSector(showTransactionModal.item.location === 'Farm√°cia' ? 'Farm√°cia' : 'Almoxarifado');
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
      doc.text('Relat√≥rio PCA - Plano Anual de Contrata√ß√£o', 105, currentY, { align: 'center' });
      currentY += 8;
      
      doc.setFontSize(10);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(120, 113, 108);
      doc.text(`Per√≠odo de Consumo: ${format(start, 'dd/MM/yyyy')} at√© ${format(end, 'dd/MM/yyyy')}`, 105, currentY, { align: 'center' });
      currentY += 5;
      doc.text(`Gerado em: ${format(new Date(), 'dd/MM/yyyy HH:mm')}`, 105, currentY, { align: 'center' });
      currentY += 15;

      const categories = Object.keys(groupedData).sort();
      
      if (categories.length === 0) {
        doc.setFontSize(12);
        doc.text('Nenhum consumo registrado no per√≠odo selecionado.', 105, currentY + 20, { align: 'center' });
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
      showToast("Relat√≥rio PCA gerado com sucesso!", "success");
    } catch (error) {
      console.error('Erro ao gerar relat√≥rio PCA:', error);
      showToast("Erro ao gerar relat√≥rio PCA.", "error");
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
          { name: 'HIGIENIZA√á√ÉO', values: [4, 0, 0, 1, 3, 0], total: 8 },
          { name: 'ILHA', values: [316, 178, 266, 310, 579, 200], total: 1849 },
          { name: 'IMAGEM', values: [351, 354, 131, 267, 505, 106], total: 1714 },
          { name: 'P√â DIAB√âTICO', values: [384, 476, 563, 548, 572, 552], total: 3095 },
          { name: 'RECEP√á√ÉO GERAL', values: [203, 0, 0, 0, 110, 17], total: 330 },
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
      months = ['Janeiro', 'Fevereiro', 'Mar√ßo', 'Abril', 'Maio', 'Junho'];
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
      months = ['Janeiro', 'Fevereiro', 'Mar√ßo', 'Abril', 'Maio', 'Junho'];
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

      if (filterLower.includes('m√©dico') || filterLower.includes('medico') || filterLower.includes('hospitalar')) {
        return catLower.includes('m√©dico') || catLower.includes('medico') || catLower.includes('hospitalar');
      }
      if (filterLower.includes('medicamento')) {
        return catLower.includes('medicamento') || catLower.includes('f√°rmaco') || catLower.includes('farmaco');
      }
      if (filterLower.includes('aliment')) {
        return catLower.includes('aliment') || catLower.includes('copa') || catLower.includes('cozinha');
      }
      if (filterLower.includes('expediente')) {
        return catLower.includes('expediente') || catLower.includes('papelaria') || catLower.includes('escrit√≥rio') || catLower.includes('escritorio');
      }
      if (filterLower.includes('higiene') || filterLower.includes('limpeza')) {
        return catLower.includes('higiene') || catLower.includes('limpeza') || catLower.includes('saneante');
      }
      if (filterLower.includes('odont')) {
        return catLower.includes('odont');
      }
      if (filterLower.includes('epi')) {
        return catLower.includes('epi') || catLower.includes('seguran√ßa') || catLower.includes('seguranca');
      }
      if (filterLower.includes('inform√°t') || filterLower.includes('informat') || filterLower.includes('ti')) {
        return catLower.includes('inform√°t') || catLower.includes('informat') || catLower.includes('ti');
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
          { name: 'P√â DIAB√âTICO', values: [384, 476, 563, 548, 572, 552], total: 3095 },
          { name: 'ILHA', values: [316, 178, 266, 310, 579, 200], total: 1849 },
          { name: 'IMAGEM', values: [351, 354, 131, 267, 505, 106], total: 1714 }
        ]);

    const activeSectorsForAnalysis = finalSectors.filter(s => s.total > 0);

    let periodText = 'no per√≠odo analisado';
    if (quantitativoPeriodPreset === '1_semestre_2026') periodText = 'no 1¬∫ semestre de 2026';
    else if (quantitativoPeriodPreset === '2_semestre_2026') periodText = 'no 2¬∫ semestre de 2026';
    else if (quantitativoPeriodPreset === 'ano_2026') periodText = 'no ano de 2026 (total)';

    const catLabel = quantitativoCategory === 'Todos' ? 'materiais e insumos em geral' : `materiais da categoria ${quantitativoCategory.toUpperCase()}`;

    let autoAnalysis = '';
    if (activeSectorsForAnalysis.length > 0) {
      const top1 = activeSectorsForAnalysis[0];
      const top2 = activeSectorsForAnalysis[1];
      const grandTotal = activeSectorsForAnalysis.reduce((acc, s) => acc + s.total, 0);

      const monthTotals = months.map((_, idx) => activeSectorsForAnalysis.reduce((sum, sec) => sum + (sec.values[idx] || 0), 0));
      const maxMonthIdx = monthTotals.indexOf(Math.max(...monthTotals));
      const maxMonthName = months[maxMonthIdx] || 'm√™s de pico';

      let sector2Text = '';
      if (top2 && top2.total > 0) {
        sector2Text = ` Em SEGUNDO LUGAR, destaca-se o setor de ${top2.name}, acumulando ${top2.total.toLocaleString('pt-BR')} unidades (${((top2.total / grandTotal) * 100).toFixed(1)}% do total).`;
      }

      autoAnalysis = `Verificou-se que, ${periodText}, o volume total de dispensa√ß√£o para ${catLabel} foi de ${grandTotal.toLocaleString('pt-BR')} unidades. O setor com MAIOR DEMANDA foi o de ${top1.name}, apresentando ${top1.total.toLocaleString('pt-BR')} unidades dispensadas (${((top1.total / grandTotal) * 100).toFixed(1)}% do consumo total).${sector2Text} Observou-se o maior pico de dispensa√ß√µes no m√™s de ${maxMonthName}. Os dados registrados pelo sistema indicam maior concentra√ß√£o assistencial nesses setores e auxiliam no planejamento das compras e estoques do almoxarifado.`;
    } else {
      autoAnalysis = `Verificou-se que, ${periodText}, n√£o foram registradas movimenta√ß√µes de sa√≠da ou solicita√ß√µes entregues para ${catLabel} no sistema. Os controles de estoque do almoxarifado permanecem monitorando o fluxo de demandas.`;
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
      showToast("Gerando PDF oficial do relat√≥rio...", "info");
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
      showToast("Erro ao gerar PDF. Tente usar a fun√ß√£o de impress√£o.", "error");
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
      doc.text(`Local F√≠sico Origem: ${roomFilter}`, 14, 46);
      doc.text(`Emitido em: ${format(new Date(), 'dd/MM/yyyy HH:mm')}`, 14, 52);

      // Filter items by room and categories
      const roomItems = items.filter(i => {
        // Ignorar exclu√≠dos ou sem estoque
        if (i.deletedAt || i.quantity <= 0) return false;
        
        // Normaliza√ß√£o para compara√ß√£o robusta
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
            else if (daysToExpiry <= 30) expiryStatus = 'CR√çTICO';
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
            
            // For√ßamos o formato JPEG com qualidade alta
            const dataUrl = canvas.toDataURL('image/jpeg', 0.95);
            resolve(dataUrl);
            URL.revokeObjectURL(img.src);
          } catch (e) {
            reject(e);
          }
        };

        img.onerror = () => {
          clearTimeout(timeout);
          // Se o Canvas falhar, tenta FileReader direto como √∫ltimo recurso
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
    items: { product_name: string; quantity: number; batch_number?: string; expiry_date?: string }[];
    revisionDate: string;
    donationNumber?: string;
    date: string;
  }) => {
    try {
      showToast("Gerando Termo de Doa√ß√£o...", "info");

      let base64Image = letterheadImage || "";
      
      // Se n√£o houver imagem personalizada, tenta carregar a padr√£o
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
            console.log("[PDF] Desenhando imagem de papel timbrado no Termo de Doa√ß√£o");
            const format = base64Image.includes('image/png') ? 'PNG' : 'JPEG';
            pdfDoc.addImage(base64Image, format, 0, 0, pageWidth, pageHeight, undefined, 'FAST');
            return;
          } catch (e) {
            console.error("Error adding letterhead image to Donation Term:", e);
          }
        }
        
        console.log("[PDF] Usando cabe√ßalho padr√£o com 3 logos retangulares expandidos no Termo de Doa√ß√£o");
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

        // 2. LOGO POLICL√çNICA (Center - Rectangular)
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
          pdfDoc.text('CEO - CENTRO DE ESPECIALIDADES ODONTOL√ìGICAS', centerX + (logoWidth / 2), logoY + 10, { align: 'center' });
        }

        // 3. LOGO CONS√ìRCIO CPSMS (Right - Rectangular)
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
          pdfDoc.text('CONS√ìRCIO CPSMS', rightX + (consorcioWidth / 2), consorcioY + 11, { align: 'center' });
        }

        pdfDoc.setDrawColor(226, 232, 240);
        pdfDoc.setLineWidth(0.5);
        pdfDoc.line(margin, 29, pageWidth - margin, 29);

        // Footer
        pdfDoc.setFontSize(7.5);
        pdfDoc.setTextColor(120, 113, 108);
        pdfDoc.setFont('helvetica', 'normal');
        const footer1 = 'CEO - Centro de Especialidades Odontol√≥gicas.';
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

      const donorName = formatTitleCase(data.donatingUnitName || 'CEO - Centro de Especialidades Odontol√≥gicas');
      const receivingName = formatTitleCase(data.receivingUnit.name);
      const receivingAddress = data.receivingUnit.address;
      const receivingCNPJ = data.receivingUnit.cnpj;

      drawLetterhead(doc);

      // --- TITLE & DATA DE EMISS√ÉO BELOW LOGOS ---
      doc.setFontSize(13);
      doc.setTextColor(17, 24, 39);
      doc.setFont('helvetica', 'bold');
      doc.text('TERMO DE DOA√á√ÉO DE MATERIAIS E INSUMOS', pageWidth / 2, 35, { align: 'center' });
      
      doc.setFontSize(8.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(100, 116, 139);
      doc.text(`Data de Emiss√£o: ${format(new Date(), 'dd/MM/yyyy HH:mm')}`, pageWidth / 2, 40, { align: 'center' });

      // --- DOCUMENT METADATA RIGHT-ALIGNED ---
      doc.setFontSize(8);
      doc.setTextColor(107, 114, 128);
      doc.setFont('helvetica', 'normal');
      doc.text('C√≥digo: TERMO-ALMOX', pageWidth - margin, 46, { align: 'right' });
      doc.text(`Data de Implanta√ß√£o: ${format(new Date(), 'dd/MM/yyyy')}`, pageWidth - margin, 50, { align: 'right' });
      doc.text(`√öltima Revis√£o: ${data.revisionDate || '---'}`, pageWidth - margin, 54, { align: 'right' });
      
      if (data.donationNumber) {
        doc.setFontSize(9);
        doc.setTextColor(31, 41, 55);
        doc.setFont('helvetica', 'bold');
        doc.text(`Termo n¬∫: ${data.donationNumber}`, pageWidth - margin, 59, { align: 'right' });
      }

      doc.setDrawColor(209, 213, 219);
      doc.setLineWidth(0.2);
      doc.line(margin, 63, pageWidth - margin, 63);

      // --- CONTENT ---
      doc.setFontSize(10);
      doc.setTextColor(31, 41, 55);
      doc.setFont('helvetica', 'normal');
      
      const donationText = `A ${donorName}, inscrita sob o CNPJ n¬∫ 12.208.466/0001-66, por interm√©dio de seu Setor de Almoxarifado, formaliza por este instrumento a doa√ß√£o √† unidade ${receivingName}, situada em ${receivingAddress}, inscrita sob o CNPJ n¬∫ ${receivingCNPJ}, dos materiais e insumos abaixo discriminados. A presente cess√£o justifica-se pela otimiza√ß√£o de estoque em virtude da redu√ß√£o de demanda interna e proximidade do prazo de validade, assegurando a destina√ß√£o √∫til dos itens.`;
      
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
        head: [['Descri√ß√£o do Material', 'Lote', 'Qtd Doada', 'Confer√™ncia']],
        body: data.items.map(i => [
          i.product_name, 
          (i.batch_number && i.batch_number.trim() !== '') ? i.batch_number : '---',
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
           1: { cellWidth: 32, halign: 'center', fontStyle: 'bold' },
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
      showToast("Termo de Doa√ß√£o gerado com sucesso!", "success");
    } catch (error) {
      console.error('Erro ao exportar PDF de Doa√ß√£o:', error);
      alert('Ocorreu um erro ao gerar o Termo de Doa√ß√£o.');
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

      // 2. CENTER LOGO: Logo da Policl√≠nica (Rectangular)
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
        doc.text('CEO - CENTRO DE ESPECIALIDADES ODONTOL√ìGICAS', centerX + (logoWidth / 2), logoY + 10, { align: 'center' });
      }

      // 3. RIGHT LOGO: Logo do Cons√≥rcio CPSMS (Rectangular)
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
        doc.text('CONS√ìRCIO CPSMS', rightX + (consorcioWidth / 2), consorcioY + 11, { align: 'center' });
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
    doc.text(`Data de Emiss√£o: ${format(new Date(), 'dd/MM/yyyy HH:mm')}`, pageWidth / 2, startY, { align: 'center' });
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
      showToast("Imagem muito grande. M√°ximo 5MB.", "error");
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
        showToast("Imagem muito grande. M√°ximo 2MB.", "error");
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
        showToast("Imagem muito grande. M√°ximo 2MB.", "error");
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
        showToast("Imagem muito grande. M√°ximo 2MB.", "error");
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
      console.error("Erro ao remover logo da Policl√≠nica do Firestore:", err);
      showToast("Logo do CEO removida!", "success");
    }
  };

  const handleConsorcioLogoUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.size > 2 * 1024 * 1024) {
        showToast("Imagem muito grande. M√°ximo 2MB.", "error");
        return;
      }
      const reader = new FileReader();
      reader.onloadend = async () => {
        const base64 = reader.result as string;
        setConsorcioLogo(base64);
        localStorage.setItem('consorcio_logo_base64', base64);
        try {
          await setDoc(doc(db, 'settings', 'general'), { consorcioLogo: base64 }, { merge: true });
          showToast("Logo do Cons√≥rcio CPSMS atualizada com sucesso!", "success");
        } catch (err) {
          console.error("Erro ao salvar no Firestore:", err);
          showToast("Logo do Cons√≥rcio atualizada!", "success");
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
      showToast("Logo do Cons√≥rcio removida com sucesso!", "success");
    } catch (err) {
      console.error("Erro ao remover logo do Cons√≥rcio do Firestore:", err);
      showToast("Logo do Cons√≥rcio removida!", "success");
    }
  };

  const handleExportDeliveryReceiptPDF = async (data: {
    sector: string;
    items: { product_name: string; quantity: number; batch_number?: string; expiry_date?: string; unit_measure?: string }[];
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

        // 2. LOGO POLICL√çNICA (Center - Rectangular)
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
          pdfDoc.text('CEO - CENTRO DE ESPECIALIDADES ODONTOL√ìGICAS', centerX + (logoWidth / 2), logoY + 10, { align: 'center' });
        }

        // 3. LOGO CONS√ìRCIO CPSMS (Right - Rectangular)
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
          pdfDoc.text('CONS√ìRCIO CPSMS', rightX + (consorcioWidth / 2), consorcioY + 11, { align: 'center' });
        }

        pdfDoc.setDrawColor(226, 232, 240);
        pdfDoc.setLineWidth(0.5);
        pdfDoc.line(margin, 29, pageWidth - margin, 29);

        // Footer
        pdfDoc.setFontSize(7.5);
        pdfDoc.setTextColor(120, 113, 108);
        pdfDoc.setFont('helvetica', 'normal');
        const footer1 = 'CEO - Centro de Especialidades Odontol√≥gicas.';
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
      doc.text(`Data de Emiss√£o: ${format(new Date(), 'dd/MM/yyyy HH:mm')}`, pageWidth / 2, 40, { align: 'center' });

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
      doc.text('REFER√äNCIA:', 19, 68);
      
      doc.setTextColor(30, 41, 59);
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(11);
      doc.text(data.sector.toUpperCase(), 52, 60);
      
      doc.setFontSize(9);
      doc.setFont('helvetica', 'normal');
      doc.text(data.requestId ? `Solicita√ß√£o #${data.requestId.slice(-5).toUpperCase()}` : 'Baixa Direta no Sistema', 52, 68);
      
      doc.text('DATA DA SA√çDA:', pageWidth - 80, 68);
      doc.setFont('helvetica', 'bold');
      doc.text(format(new Date(data.date), 'dd/MM/yyyy'), pageWidth - 50, 68);

      // Materials Table
      const tableData = data.items.map(i => [
        i.product_name.toUpperCase(), 
        (i.batch_number && i.batch_number.trim() !== '') ? i.batch_number.toUpperCase() : '---',
        i.quantity.toString(), 
        '_________________'
      ]);
      
      autoTable(doc, {
        startY: 80,
        head: [['DESCRI√á√ÉO DO MATERIAL', 'LOTE', 'QTD ENTREGUE', 'CONFER√äNCIA']],
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
           1: { cellWidth: 32, halign: 'center', fontStyle: 'bold' },
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
      doc.text('RESPONS√ÅVEL PELA ENTREGA', 20 + (signLineW/2), finalY + 5, { align: 'center' });
      doc.text('RESPONS√ÅVEL PELO SETOR (RECEBIMENTO)', pageWidth - 20 - (signLineW/2), finalY + 5, { align: 'center' });
      
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(7);
      const responsibleName = userProfile?.name || user?.displayName || 'Respons√°vel';
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
      
      const footerLine1 = 'CEO - Centro de Especialidades Odontol√≥gicas.';
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
      doc.text('GEST√ÉO DE ALMOXARIFADO E FARM√ÅCIA', 28, 22);

      doc.setDrawColor(231, 229, 228); // light border
      doc.setLineWidth(0.5);
      doc.line(14, 28, pageWidth - 14, 28);
      
      // Title and Date
      doc.setFontSize(14);
      doc.setTextColor(28, 25, 23);
      doc.setFont('helvetica', 'bold');
      doc.text('Relat√≥rio de Consumo por Setor', 14, 40);
      
      doc.setFontSize(9);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(120, 113, 108);
      doc.text(`Per√≠odo: ${format(parseISO(reportRange.start), 'dd/MM/yyyy')} a ${format(parseISO(reportRange.end), 'dd/MM/yyyy')}`, 14, 46);
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
          doc.text(`Documento emitido pelo Sistema de Gest√£o Hospitalar - P√°gina ${doc.getNumberOfPages()}`, 14, doc.internal.pageSize.height - 10);
        }
      });
      
      const fileName = `Relatorio_Consumo_CEO_${format(new Date(), 'dd-MM-yyyy')}.pdf`;
      doc.save(fileName);
      showToast("Relat√≥rio profissional exportado!", "success");
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
                          (effectiveSectorFilter === 'Farm√°cia' && t.sector === 'Farm√°cia (Consumo Interno)');
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
        const userLocation = userProfile?.sector === 'Farm√°cia' ? 'Farm√°cia' : 'Almoxarifado';
        return (item.location || 'Almoxarifado') === userLocation;
      }
      
      // If admin, respect the sector filter if it maps to a location
      if (reportSectorFilter === 'Farm√°cia') {
        return item.location === 'Farm√°cia';
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
        const userLocation = userProfile?.sector === 'Farm√°cia' ? 'Farm√°cia' : 'Almoxarifado';
        return (item.location || 'Almoxarifado') === userLocation;
      }
      
      // If admin, respect the sector filter if it maps to a location
      if (reportSectorFilter === 'Farm√°cia') {
        return item.location === 'Farm√°cia';
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
      
      const sectorKey = (t.sector === 'Farm√°cia (Consumo Interno)') ? 'Farm√°cia' : t.sector!;
      
      if (!sectorData[sectorKey]) {
        sectorData[sectorKey] = { name: sectorKey };
      }
      sectorData[sectorKey][category] = (sectorData[sectorKey][category] || 0) + t.quantity;
    });

    filteredTrans.filter(t => t.type === 'entry' && t.isReturn && t.sector).forEach(t => {
      const item = items.find(i => i.id === t.item_id);
      const category = item?.category || 'Outros';
      categoriesInSector.add(category);
      
      const sectorKey = (t.sector === 'Farm√°cia (Consumo Interno)') ? 'Farm√°cia' : t.sector!;
      
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
      let sector = t.sector || 'N√£o Informado';
      if (sector === 'Farm√°cia (Consumo Interno)') sector = 'Farm√°cia';
      
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
      let sector = t.sector || 'N√£o Informado';
      if (sector === 'Farm√°cia (Consumo Interno)') sector = 'Farm√°cia';
      
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
      let sec = t.sector || 'N√£o Informado';
      if (sec === 'Farm√°cia (Consumo Interno)') sec = 'Farm√°cia';

      if (!returnsBySectorMap[sec]) {
        returnsBySectorMap[sec] = { name: sec, quantity: 0, value: 0 };
      }
      returnsBySectorMap[sec].quantity += t.quantity;
      returnsBySectorMap[sec].value += val;

      const reason = t.returnReason || 'N√£o especificado';
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
            <p className="text-[10px] text-slate-400 uppercase tracking-widest font-extrabold">Acesso restrito a funcion√°rios autorizados</p>
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
                      <span>Hist√≥rico</span>
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
                      <span>Solicita√ß√µes</span>
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
                      <span>Devolu√ß√µes</span>
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
                      <span>Relat√≥rios</span>
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
                      <span>Usu√°rios</span>
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
                      <span>{selectedSector === 'Farm√°cia' || userProfile?.sector === 'Farm√°cia' ? 'Estoque da Farm√°cia' : 'Estoque'}</span>
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
                      <span>Nova Solicita√ß√£o</span>
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
                      <span>Devolu√ß√£o de Materiais</span>
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
                      <span>Minhas Solicita√ß√µes</span>
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
                      <span>Relat√≥rios</span>
                    </div>
                  </button>

                  {(userProfile?.role === 'L√çDER' || userProfile?.role === 'SETOR') && (
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
                        <span>Estat√≠sticas</span>
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
            <p className="text-[10px] font-black text-slate-400 uppercase tracking-wider mb-2">Usu√°rio Conectado</p>
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
            <Settings size={18} className="text-slate-400" /> Configura√ß√µes
          </button>
        </div>
      </aside>

      {/* Main Content */}
      <main className="lg:ml-64 p-4 lg:p-10 max-w-7xl mx-auto mt-16 lg:mt-0">
        <header className="flex flex-col lg:flex-row lg:justify-between lg:items-center gap-4 mb-6 lg:mb-10">
          <div>
            <h2 className="text-xl lg:text-3xl font-bold tracking-tight mb-1">
              {activeTab === 'dashboard' && 'Vis√£o Geral'}
              {activeTab === 'inventory' && 'Gerenciamento de Estoque'}
              {activeTab === 'history' && 'Hist√≥rico de Movimenta√ß√µes'}
              {activeTab === 'requests' && 'Solicita√ß√µes de Materiais'}
              {activeTab === 'admin-devolutions' && 'Devolu√ß√µes de Materiais'}
              {activeTab === 'trash' && 'Lixeira (Exclus√£o em 3 dias)'}
              {activeTab === 'my-requests' && `Minhas Solicita√ß√µes - ${selectedSector || ''}`}
              {activeTab === 'new-request' && `Nova Solicita√ß√£o - ${selectedSector || ''}`}
              {activeTab === 'devolution' && `Devolu√ß√£o de Materiais - ${selectedSector || ''}`}
              {editingRequest && ' - Editando Solicita√ß√£o'}
              {activeTab === 'reports' && 'Relat√≥rios e An√°lises'}
              {activeTab === 'leader-stats' && 'Estat√≠sticas do Almoxarifado'}
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
                        onChange={e => setInventoryLocation(e.target.value as 'Almoxarifado' | 'Farm√°cia')}
                      >
                        <option value="Almoxarifado">Almoxarifado</option>
                        <option value="Farm√°cia">Farm√°cia</option>
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
                    <span className="text-slate-400 text-xs font-bold">at√©</span>
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
                      <FileText size={15} /> Cat√°logo de Itens
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
                title="Notifica√ß√µes"
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
                        <h3 className="font-black text-sm">Notifica√ß√µes</h3>
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
                            <p className="text-xs text-[#A8A29E] font-medium">Nenhuma notifica√ß√£o</p>
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
                                          Confirmar Ci√™ncia
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
                    <option value="duration_asc">Dura√ß√£o (Menor-Maior)</option>
                    <option value="duration_desc">Dura√ß√£o (Maior-Menor)</option>
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
                    <option value="donation">Doa√ß√£o</option>
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
                      title="Imprimir Relat√≥rio de Itens Cr√≠ticos / Estoque Baixo"
                    >
                      <Printer size={16} className="text-amber-600" />
                      <span className="hidden sm:inline">Relat√≥rio Estoque Baixo</span>
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
                      title="Baixar Relat√≥rio PDF de Todo Estoque"
                    >
                      <Printer size={18} />
                    </button>
                  </div>
                )}
              </div>
            )}
            {activeTab === 'dashboard' && (isAdmin || selectedSector === 'Farm√°cia') && (
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
                  <ArrowUpRight size={18} /> Nova Sa√≠da
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

                {/* Card 2: Patrim√¥nio */}
                {(isAdmin || selectedSector === 'Farm√°cia') && (
                  <div className="bg-white rounded-xl border border-indigo-100/80 shadow-xs hover:shadow-sm hover:border-indigo-200 transition-all duration-200 overflow-hidden group relative">
                    <div className="h-1 w-full bg-gradient-to-r from-indigo-600 to-blue-600" />
                    <div className="p-3.5">
                      <div className="flex items-center justify-between mb-2">
                        <span className="text-[10px] font-black text-slate-500 uppercase tracking-wider">Patrim√¥nio Investido</span>
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

                {/* Card 3: Pend√™ncias / Solicita√ß√µes */}
                <div 
                  onClick={() => setActiveTab('requests')}
                  className="bg-white rounded-xl border border-sky-100/80 shadow-xs hover:shadow-sm hover:border-sky-300 transition-all duration-200 overflow-hidden group cursor-pointer relative"
                >
                  <div className="h-1 w-full bg-gradient-to-r from-sky-500 to-blue-600" />
                  <div className="p-3.5">
                    <div className="flex items-center justify-between mb-2">
                      <span className="text-[10px] font-black text-slate-500 uppercase tracking-wider">Solicita√ß√µes Pendentes</span>
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
                        {pendingRequestsCount > 0 ? 'Aguardando atendimento' : 'Nenhuma pend√™ncia'}
                      </span>
                    </div>
                  </div>
                </div>

                {/* Card 4: N√≠vel Cr√≠tico / Alertas */}
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
                      <span className="text-[10px] font-black text-slate-500 uppercase tracking-wider">Aten√ß√£o Necess√°ria</span>
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
                              {nearExpiryItems.length} pr√≥x. vencer
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
                          Distribui√ß√£o de Estoque por Categoria
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
                {/* Left Column: Central de Alertas Cr√≠ticos (5/12 cols) */}
                <div className="lg:col-span-5 bg-white rounded-3xl border border-blue-100 shadow-sm overflow-hidden space-y-0">
                  <div className="bg-gradient-to-r from-slate-900 via-blue-950 to-slate-900 text-white p-5 flex items-center justify-between">
                    <div className="flex items-center gap-2.5">
                      <div className="p-2 bg-amber-500/20 text-amber-400 rounded-xl border border-amber-400/30">
                        <AlertTriangle size={18} />
                      </div>
                      <div>
                        <h4 className="text-xs font-extrabold uppercase tracking-wider text-white">
                          Central de Alertas Cr√≠ticos
                        </h4>
                        <p className="text-[10px] text-blue-200 font-medium">
                          Itens vencidos, com estoque baixo ou pr√≥ximos ao vencimento
                        </p>
                      </div>
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={handleExportLowStockPDF}
                        className="px-2.5 py-1 bg-amber-500/20 hover:bg-amber-500/30 text-amber-200 border border-amber-400/40 rounded-xl text-[11px] font-extrabold transition-all flex items-center gap-1 shadow-xs"
                        title="Imprimir Relat√≥rio PDF dos Itens Cr√≠ticos / Estoque Baixo"
                      >
                        <Printer size={13} /> Relat√≥rio PDF
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
                          Nenhum insumo apresentou n√≠vel cr√≠tico de reposi√ß√£o, vencimento ultrapassado ou data de expira√ß√£o pr√≥xima.
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
                              Abaixo do m√≠nimo recomendado ({group.min_quantity} un)
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
                              <span className="text-[9px] font-black uppercase px-1.5 py-0.2 rounded bg-sky-200 text-sky-900">Pr√≥x. Vencer</span>
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

                {/* Right Column: Movimenta√ß√µes Recentes (7/12 cols) */}
                <div className="lg:col-span-7 bg-white rounded-3xl border border-blue-100 shadow-sm overflow-hidden space-y-0">
                  <div className="bg-gradient-to-r from-slate-900 via-blue-950 to-slate-900 text-white p-5 flex items-center justify-between">
                    <div className="flex items-center gap-2.5">
                      <div className="p-2 bg-blue-500/20 text-blue-300 rounded-xl border border-blue-400/30">
                        <History size={18} />
                      </div>
                      <div>
                        <h4 className="text-xs font-extrabold uppercase tracking-wider text-white">
                          Movimenta√ß√µes Recentes do Estoque
                        </h4>
                        <p className="text-[10px] text-blue-200 font-medium">
                          Hist√≥rico de sa√≠das e entradas registradas no almoxarifado
                        </p>
                      </div>
                    </div>
                    <span className="text-[10px] font-bold uppercase tracking-wider bg-blue-500/20 text-blue-300 border border-blue-400/30 px-2.5 py-1 rounded-full">
                      √öltimos 5 registros
                    </span>
                  </div>

                  <div className="p-5 space-y-3 max-h-[480px] overflow-y-auto">
                    {recentTransactions.length === 0 && (
                      <div className="py-14 text-center">
                        <History size={40} className="mx-auto text-slate-300 mb-3" />
                        <p className="text-sm text-slate-800 font-bold">Sem registros no momento</p>
                        <p className="text-xs text-slate-500">Nenhuma movimenta√ß√£o realizada nesta localiza√ß√£o.</p>
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
                            {t.type === 'entry' ? 'Entrada / Adi√ß√£o em estoque' : `Sa√≠da e entrega p/ setor: ${t.sector || '---'}`}
                          </p>
                          <div className="flex flex-wrap items-center gap-x-2 gap-y-1 mt-2 pt-1.5 border-t border-dashed border-slate-200/80">
                            <span className="text-[10px] text-slate-400 font-medium flex items-center gap-1">
                              <Clock size={11} />
                              {new Date(t.date).toLocaleDateString('pt-BR')} √†s {new Date(t.date).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}
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
                            title="Excluir Movimenta√ß√£o"
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
                      onClick={() => setInventoryLocation('Farm√°cia')}
                      className={`px-5 py-2 rounded-xl text-xs font-extrabold transition-all duration-200 flex items-center gap-2 ${
                        inventoryLocation === 'Farm√°cia' 
                          ? 'bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white shadow-md shadow-blue-600/20' 
                          : 'text-slate-600 hover:bg-slate-200/70'
                      }`}
                    >
                      <Users size={15} /> Estoque Farm√°cia
                    </button>
                  </div>
                ) : (
                  <div className="flex items-center gap-3 px-4 py-2 bg-blue-50/80 rounded-2xl border border-blue-100">
                    <div className="p-2 bg-gradient-to-br from-blue-700 to-indigo-900 text-white rounded-xl shadow-sm">
                      {inventoryLocation === 'Farm√°cia' ? <Users size={16} /> : <Package size={16} />}
                    </div>
                    <div>
                      <p className="text-xs font-black text-slate-900">
                        Estoque: <span className="text-blue-700">{inventoryLocation === 'Farm√°cia' ? 'Medicamentos (Farm√°cia)' : 'Almoxarifado Geral'}</span>
                      </p>
                      <p className="text-[10px] text-slate-500 font-medium">Acesso exclusivo aos medicamentos da Farm√°cia</p>
                    </div>
                  </div>
                )}

                <div className="flex items-center gap-3">
                  {inventoryLocation === 'Farm√°cia' && (
                    <button 
                      onClick={() => setActiveTab('new-request')}
                      className="px-4 py-2.5 bg-gradient-to-r from-emerald-600 to-teal-700 hover:from-emerald-700 hover:to-teal-800 text-white font-extrabold text-xs rounded-2xl shadow-md hover:shadow-lg transition-all flex items-center gap-2"
                      title="Solicitar novos medicamentos ao Almoxarifado Geral"
                    >
                      <Plus size={16} /> Solicitar ao Almoxarifado
                    </button>
                  )}
                  <span className="text-xs font-black text-slate-500 uppercase tracking-wider bg-slate-50 border border-slate-200 px-3 py-1.5 rounded-xl">
                    Visualiza√ß√£o: <span className="text-blue-700 font-black">{inventoryLocation}</span>
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
                    <th className="px-6 py-4 font-black text-xs text-blue-200/90 uppercase tracking-wider">{isAdmin ? 'Pre√ßo Un.' : '---'}</th>
                    <th className="px-6 py-4 font-black text-xs text-blue-200/90 uppercase tracking-wider text-center">Quantidade</th>
                    <th className="px-6 py-4 font-black text-xs text-blue-200/90 uppercase tracking-wider">M√≠nimo</th>
                    <th className="px-6 py-4 font-black text-xs text-blue-200/90 uppercase tracking-wider text-center">Dura√ß√£o</th>
                    <th className="px-6 py-4 font-black text-xs text-blue-200/90 uppercase tracking-wider">Status Cr√≠tico</th>
                    <th className="px-6 py-4 font-black text-xs text-blue-200/90 uppercase tracking-wider text-right">A√ß√µes</th>
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
                                    {origin === 'contract' ? 'Contrato' : origin === 'donation' ? 'Doa√ß√£o' : 'Extra'}
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
                              {group.durationWeeks === 'infinite' ? '‚àû' : `${group.durationWeeks.toFixed(1)} sem`}
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
                              {group.durationWeeks <= 4 ? 'Muito Cr√≠tico' :
                               group.durationWeeks <= 8 ? 'Aten√ß√£o' :
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
                                {item.origin === 'contract' ? 'Contrato' : item.origin === 'donation' ? 'Doa√ß√£o' : 'Extra'}
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
                                    title="Editar Pre√ßo"
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
                              title="Registrar Sa√≠da"
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
                  <h3 className="text-lg font-bold text-[#1C1917]">Hist√≥rico de Movimenta√ß√µes</h3>
                  {isAdmin && (
                    <div className="flex items-center gap-2 bg-[#F5F5F4] p-1 rounded-2xl border border-[#E7E5E4]">
                      <button 
                        onClick={() => setInventoryLocation('Almoxarifado')}
                        className={`px-4 py-1.5 rounded-xl text-[10px] font-bold transition-all ${inventoryLocation === 'Almoxarifado' ? 'bg-[#1C1917] text-white shadow-sm' : 'text-[#78716C] hover:bg-[#E7E5E4]'}`}
                      >
                        Almoxarifado
                      </button>
                      <button 
                        onClick={() => setInventoryLocation('Farm√°cia')}
                        className={`px-4 py-1.5 rounded-xl text-[10px] font-bold transition-all ${inventoryLocation === 'Farm√°cia' ? 'bg-[#1C1917] text-white shadow-sm' : 'text-[#78716C] hover:bg-[#E7E5E4]'}`}
                      >
                        Farm√°cia
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
                    {showDeletedHistory ? 'Ver Hist√≥rico Ativo' : 'Ver Exclu√≠dos (Testes)'}
                  </button>
                </div>
              </div>

              <div className="bg-white rounded-3xl border border-[#E7E5E4] shadow-sm overflow-hidden">
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse min-w-[1200px]">
                  <thead>
                    <tr className="bg-[#FAFAF9] border-bottom border-[#E7E5E4]">
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Data</th>
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Movimenta√ß√£o</th>
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Item</th>
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Lote</th>
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Validade</th>
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-center">Origem</th>
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Setor</th>
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">Respons√°vel</th>
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">Qtd</th>
                    {isAdmin && <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right whitespace-nowrap">Val. Unit</th>}
                    {isAdmin && <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right whitespace-nowrap">Total</th>}
                    <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">A√ß√µes</th>
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
                            {t.type === 'entry' ? 'Entrada' : 'Sa√≠da'}
                          </span>
                        </td>
                        <td className="px-6 py-5">
                          <div className="font-bold whitespace-nowrap">{t.item_name}</div>
                          {t.exitReason && t.exitReason !== 'consumo' && (
                            <div className="text-[10px] text-rose-500 font-bold mt-1 uppercase">
                              Motivo: {t.exitReason === 'vencido' ? 'Vencimento' : t.exitReason === 'doacao' ? 'Doa√ß√£o' : t.exitReason === 'perda' ? 'Perda/Avaria' : t.exitReason}
                              {t.expiryReason && <span className="text-[#78716C] lowercase font-normal ml-1">({t.expiryReason})</span>}
                            </div>
                          )}
                          {t.deletionReason && (
                            <div className="text-[10px] text-rose-500 font-bold mt-1">Exclus√£o: {t.deletionReason}</div>
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
                            {t.origin === 'contract' ? 'Contrato' : t.origin === 'donation' ? 'Doa√ß√£o' : 'Extra'}
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
                                      items: [{ 
                                        product_name: t.item_name, 
                                        quantity: t.quantity,
                                        batch_number: t.batch_number || '',
                                        expiry_date: t.expiry_date || ''
                                      }],
                                      revisionDate: t.donationRevisionDate || '',
                                      donationNumber: t.donationNumber,
                                      date: t.date
                                    });
                                  } else {
                                    handleExportDeliveryReceiptPDF({
                                      sector: t.sector || 'Sem Setor',
                                      items: [{ 
                                        product_name: t.item_name, 
                                        quantity: t.quantity,
                                        batch_number: t.batch_number || '',
                                        expiry_date: t.expiry_date || ''
                                      }],
                                      date: t.date
                                    });
                                  }
                                }}
                                className="p-2 text-blue-600 hover:bg-blue-50 rounded-lg transition-all"
                                title={t.exitReason === 'doacao' ? 'Reimprimir Termo de Doa√ß√£o' : 'Reimprimir Recibo de Entrega'}
                              >
                                {t.exitReason === 'doacao' ? <FileText size={18} /> : <Printer size={18} />}
                              </button>
                            )}
                            {t.deletedAt ? (
                              <button 
                                onClick={() => handleRecoverTransaction(t.id)}
                                className="p-2 text-emerald-600 hover:bg-emerald-50 rounded-lg transition-all"
                                title="Recuperar Movimenta√ß√£o"
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
                                title="Apagar Movimenta√ß√£o"
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
                    <p className="text-[#78716C]">Nenhuma movimenta√ß√£o encontrada para {inventoryLocation}.</p>
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
                        Intelig√™ncia Anal√≠tica de Estoque
                      </span>
                      <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-slate-100 text-slate-600 border border-slate-200/80">
                        {reportSectorFilter === 'all' ? 'Todos os Setores' : reportSectorFilter}
                      </span>
                    </div>

                    <h2 className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">
                      Relat√≥rios & Gest√£o de Consumo
                    </h2>
                    <p className="text-xs sm:text-sm text-slate-500 font-medium leading-relaxed">
                      Proje√ß√µes or√ßament√°rias, hist√≥rico de sa√≠das, curva de movimenta√ß√£o f√≠sica e relat√≥rios fiscais do almoxarifado.
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
                      <p className="text-[10px] uppercase font-extrabold text-rose-600 tracking-wider">Sa√≠das</p>
                      <p className="text-lg font-black text-slate-900 mt-0.5">{reportData.exits}</p>
                    </div>
                    <div className="h-8 w-px bg-slate-200" />
                    <div className="px-3 py-1 text-center">
                      <p className="text-[10px] uppercase font-extrabold text-amber-600 tracking-wider">Devolu√ß√µes</p>
                      <p className="text-lg font-black text-slate-900 mt-0.5">{reportData.totalReturnsCount}</p>
                    </div>
                    <div className="h-8 w-px bg-slate-200" />
                    <div className="px-3 py-1 text-center">
                      <p className="text-[10px] uppercase font-extrabold text-slate-500 tracking-wider">Per√≠odo</p>
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
                  <BarChart3 size={17} /> Relat√≥rios & Gr√°ficos
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
                          <h3 className="text-base font-black text-slate-900">Impress√£o de Solicita√ß√µes</h3>
                          <p className="text-slate-500 text-xs font-medium">Imprima as solicita√ß√µes pendentes e em separa√ß√£o por per√≠odo</p>
                        </div>
                      </div>
                    </div>

                    <div className="flex flex-col sm:flex-row items-end gap-3 mt-5 pt-4 border-t border-slate-100">
                      <div className="grid grid-cols-2 gap-3 w-full">
                        <div>
                          <label className="block text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1 ml-1">In√≠cio</label>
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
                        <Printer size={15} /> Imprimir Relat√≥rio
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
                          <h3 className="text-base font-black text-slate-900">Relat√≥rio PCA</h3>
                          <p className="text-slate-500 text-xs font-medium">Plano Anual de Contrata√ß√£o - Consumo por tipo no per√≠odo</p>
                        </div>
                      </div>
                    </div>

                    <div className="flex flex-col sm:flex-row items-end gap-3 mt-5 pt-4 border-t border-slate-100">
                      <div className="grid grid-cols-3 gap-2 w-full">
                        <div>
                          <label className="block text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1 ml-1">In√≠cio</label>
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
                          <h3 className="text-base font-black text-slate-900">D√∫vidas sobre o que pedir?</h3>
                          <p className="text-slate-500 text-xs font-medium">Baixe o cat√°logo simplificado contendo todos os nomes dos materiais e categorias cadastradas.</p>
                        </div>
                      </div>
                      <button 
                        onClick={handleExportMaterialsCatalogPDF}
                        className="bg-gradient-to-r from-blue-700 via-blue-800 to-indigo-900 text-white px-6 py-3 rounded-xl font-extrabold text-xs flex items-center justify-center gap-2 hover:from-blue-800 hover:to-indigo-950 transition-all shadow-md shadow-blue-600/20 whitespace-nowrap"
                      >
                        <Printer size={16} /> Ver Cat√°logo de Itens
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
                          <span className="text-xs font-black text-slate-500 uppercase tracking-widest">Entradas no Per√≠odo</span>
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

                    {/* Card 2: Sa√≠das */}
                    <div className="bg-white rounded-2xl border border-rose-100/80 shadow-sm hover:shadow-md hover:border-rose-200 transition-all duration-300 overflow-hidden group relative">
                      <div className="h-1.5 w-full bg-gradient-to-r from-rose-500 to-pink-600" />
                      <div className="p-6">
                        <div className="flex items-center justify-between mb-3">
                          <span className="text-xs font-black text-slate-500 uppercase tracking-widest">Sa√≠das / Consumo</span>
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
                          <span className="text-xs font-black text-slate-500 uppercase tracking-widest">Patrim√¥nio em Saldo</span>
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
                            Itens no cat√°logo
                          </span>
                        </div>
                      </div>
                    </div>
                  </>
                ) : (
                  <>
                    <div className="bg-white p-6 rounded-2xl border border-rose-100 shadow-sm hover:shadow-md transition-all lg:col-span-2 overflow-hidden relative group">
                      <div className="h-1.5 w-full bg-gradient-to-r from-rose-500 to-pink-600 absolute top-0 left-0" />
                      <p className="text-slate-500 text-xs font-black uppercase tracking-wider mb-3">Consumo do Setor no Per√≠odo</p>
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
                      <p className="text-slate-500 text-xs font-black uppercase tracking-wider mb-3">Solicita√ß√µes no Per√≠odo</p>
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
                    <h3 className="text-xl font-black text-slate-900 tracking-tight">Panorama Visual de Consumo & Movimenta√ß√£o</h3>
                    <p className="text-xs text-slate-500 font-medium">Gr√°ficos interativos para acompanhamento gerencial das opera√ß√µes</p>
                  </div>
                </div>
              </div>

              {/* Charts Grid */}
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                {/* Movement Chart */}
                <div className="bg-white p-6 sm:p-8 rounded-3xl border border-blue-100/80 shadow-sm hover:shadow-md transition-all">
                  <div className="flex items-center justify-between mb-6">
                    <h4 className="text-base font-black text-slate-900 flex items-center gap-2">
                      <Activity size={18} className="text-blue-600" /> Movimenta√ß√£o {isAdmin ? 'Geral' : 'do Setor'}
                    </h4>
                    <span className="text-[10px] font-bold text-slate-400 bg-slate-100 px-2.5 py-1 rounded-lg uppercase tracking-wider">Fluxo Di√°rio</span>
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
                        <Area type="monotone" dataKey="exits" name={isAdmin ? "Sa√≠das" : "Consumo"} stroke="#f43f5e" fillOpacity={1} fill="url(#colorExits)" strokeWidth={3} />
                      </AreaChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                {/* Category Breakdown */}
                <div className="bg-white p-6 sm:p-8 rounded-3xl border border-blue-100/80 shadow-sm hover:shadow-md transition-all">
                  <div className="flex items-center justify-between mb-6">
                    <h4 className="text-base font-black text-slate-900 flex items-center gap-2">
                      <PieChartIcon size={18} className="text-amber-500" /> Distribui√ß√£o de Consumo por Categoria
                    </h4>
                    <span className="text-[10px] font-bold text-slate-400 bg-slate-100 px-2.5 py-1 rounded-lg uppercase tracking-wider">Propor√ß√£o</span>
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
                          <TrendingDown size={18} className="text-rose-600" /> Sa√≠das por Motivo
                        </h4>
                        <span className="text-[10px] font-bold text-slate-400 bg-slate-100 px-2.5 py-1 rounded-lg uppercase tracking-wider">Destina√ß√£o</span>
                      </div>
                      <div className="h-[300px] w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <PieChart>
                            <Pie
                              data={[
                                { name: 'Consumo', value: reportData.exitsByReason.consumo },
                                { name: 'Doa√ß√£o', value: reportData.exitsByReason.doacao },
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
                        <ArrowUpRight size={18} className="text-rose-600" /> Sa√≠das por Setor (Quantidade por Tipo)
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

                {/* Returns by Sector - Devolu√ß√µes por Setor */}
                <div className="bg-white p-6 sm:p-8 rounded-3xl border border-amber-100/80 shadow-sm hover:shadow-md transition-all">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-6">
                    <div>
                      <h4 className="text-base font-black text-slate-900 flex items-center gap-2">
                        <RotateCcw size={18} className="text-amber-600" /> Devolu√ß√µes por Setor
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
                      <p className="text-sm font-black text-slate-800">Nenhuma devolu√ß√£o registrada no per√≠odo</p>
                      <p className="text-xs text-slate-500 max-w-sm mt-1 font-medium">Os materiais que forem devolvidos pelos setores ao almoxarifado no per√≠odo selecionado aparecer√£o consolidados neste gr√°fico.</p>
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
                        <BarChart3 size={18} className="text-indigo-600" /> Comparativo: Contrato vs Extra vs Doa√ß√£o
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
                              name: 'Sa√≠das', 
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
                          <Bar dataKey="doacao" name="Doa√ß√£o" fill="#10b981" radius={[6, 6, 0, 0]} barSize={26} />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                )}
              </div>

              {/* Detailed Sector Breakdown - Visible for Admin and Sector Leaders */}
              {(isAdmin || userProfile?.role === 'SETOR' || userProfile?.role === 'L√çDER') && (
                <div className="bg-white p-6 sm:p-8 rounded-3xl border border-blue-100/80 shadow-sm lg:col-span-2">
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-8 pb-6 border-b border-slate-100">
                    <div>
                      <h4 className="text-lg font-black text-slate-900 flex items-center gap-2 mb-1">
                        <History size={20} className="text-blue-600" /> 
                        Relat√≥rio Detalhado de Consumo por Item
                      </h4>
                      <p className="text-xs text-slate-500 font-medium">
                        {isAdmin ? (reportSectorFilter === 'all' ? 'Todos os Setores' : `Setor: ${reportSectorFilter}`) : `Setor: ${selectedSector}`} ‚Ä¢ {format(parseISO(reportRange.start), 'dd/MM/yyyy')} a {format(parseISO(reportRange.end), 'dd/MM/yyyy')}
                      </p>
                    </div>
                    <div className="flex flex-wrap items-center gap-3">
                      {isAdmin && (
                        <div className="text-right mr-2 hidden sm:block">
                          <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Total de Sa√≠das</p>
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
                              Nenhuma sa√≠da registrada para este per√≠odo ou setor selecionado.
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
                            Relat√≥rio Oficial Dispensa√ß√£o
                          </span>
                          <span className="px-3 py-1 rounded-full text-xs font-extrabold bg-emerald-50 text-emerald-700 border border-emerald-200">
                            {quantitativoSource === 'sample' ? 'Exemplo Oficial Sobral' : 'Dados do Sistema'}
                          </span>
                        </div>
                        <h3 className="text-xl sm:text-2xl font-black text-slate-900 mt-2">
                          Quantitativo de Materiais por Setor
                        </h3>
                        <p className="text-xs sm:text-sm text-slate-500 font-medium mt-1">
                          Gere o documento oficial com gr√°fico e an√°lise cr√≠tica para apresenta√ß√£o gerencial e fiscal referente √† categoria selecionada.
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
                          <Edit2 size={15} /> {isEditingQuantitativoAnalysis ? 'Concluir Edi√ß√£o' : 'Editar An√°lise Cr√≠tica'}
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
                          <option value="system">Dados Reais do Sistema (Padr√£o)</option>
                          <option value="sample">Exemplo Demonstrativo</option>
                        </select>
                      </div>

                      <div>
                        <label className="block text-[10px] font-black text-slate-500 uppercase tracking-wider mb-1">
                          Per√≠odo de Refer√™ncia
                        </label>
                        <select
                          value={quantitativoPeriodPreset}
                          onChange={(e) => setQuantitativoPeriodPreset(e.target.value as any)}
                          className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-bold text-xs text-slate-800 focus:ring-2 focus:ring-blue-500/20"
                        >
                          <option value="1_semestre_2026">1¬∫ Semestre de 2026 (Jan - Jun)</option>
                          <option value="2_semestre_2026">2¬∫ Semestre de 2026 (Jul - Dez)</option>
                          <option value="ano_2026">Ano Completo de 2026 (Total)</option>
                          <option value="custom">Per√≠odo Personalizado</option>
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
                          <option value="Material M√©dico-Hospitalar">Material M√©dico-Hospitalar</option>
                          <option value="Medicamentos">Medicamentos</option>
                          <option value="Aliment√≠cio">Aliment√≠cio</option>
                          <option value="Expediente">Expediente / Papelaria</option>
                          <option value="Higiene e Limpeza">Higiene e Limpeza</option>
                          <option value="Odontol√≥gico">Odontol√≥gico</option>
                          <option value="Radiol√≥gico">Radiol√≥gico</option>
                          <option value="EPI e Seguran√ßa">EPI e Seguran√ßa</option>
                          <option value="Inform√°tica">Inform√°tica / TI</option>
                          <option value="Copa & Cozinha">Copa & Cozinha</option>
                          <option value="Manuten√ß√£o">Manuten√ß√£o</option>
                          <option value="Todos">Todos os Materiais (Total Geral)</option>
                        </select>
                      </div>

                      <div>
                        <label className="block text-[10px] font-black text-slate-500 uppercase tracking-wider mb-1">
                          T√≠tulo do Documento
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
                            Texto da An√°lise Cr√≠tica (Gerada pelo Gr√°fico / Edit√°vel)
                          </label>
                          <button
                            type="button"
                            onClick={() => setQuantitativoCriticalAnalysis('')}
                            className="text-[10px] font-bold text-blue-700 hover:underline cursor-pointer flex items-center gap-1"
                          >
                            <RotateCcw size={10} /> Recalcular Autom√°tico pelo Gr√°fico
                          </button>
                        </div>
                        <textarea
                          rows={3}
                          value={quantitativoCriticalAnalysis !== '' ? quantitativoCriticalAnalysis : quantitativoReportData.criticalAnalysis}
                          onChange={(e) => setQuantitativoCriticalAnalysis(e.target.value)}
                          placeholder="Digite ou edite o texto da An√°lise Cr√≠tica do relat√≥rio..."
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

                      {/* An√°lise Cr√≠tica Section */}
                      <div className="pt-3 border-t-2 border-slate-200 space-y-2">
                        <div className="flex items-center justify-between">
                          <h3 className="text-xs font-black uppercase tracking-wider flex items-center gap-2" style={{ color: '#0f172a' }}>
                            <BarChart3 size={15} style={{ color: '#334155' }} />
                            An√°lise Cr√≠tica:
                          </h3>
                          <button
                            data-pdf-hide="true"
                            type="button"
                            onClick={() => setIsEditingQuantitativoAnalysis(!isEditingQuantitativoAnalysis)}
                            className="text-[10px] font-bold text-slate-700 hover:text-slate-900 flex items-center gap-1 cursor-pointer bg-slate-100 hover:bg-slate-200 px-2.5 py-1 rounded-lg border border-slate-300 transition-all print:hidden"
                          >
                            <Edit2 size={12} />
                            {isEditingQuantitativoAnalysis ? 'Salvar Edi√ß√£o' : 'Editar An√°lise'}
                          </button>
                        </div>

                        {isEditingQuantitativoAnalysis ? (
                          <div className="space-y-2">
                            <textarea
                              rows={6}
                              value={quantitativoCriticalAnalysis !== '' ? quantitativoCriticalAnalysis : quantitativoReportData.criticalAnalysis}
                              onChange={(e) => setQuantitativoCriticalAnalysis(e.target.value)}
                              placeholder="Digite ou edite o texto da An√°lise Cr√≠tica..."
                              className="w-full p-3 border border-slate-300 rounded-xl text-xs font-medium leading-relaxed focus:ring-2 focus:ring-slate-400/20"
                              style={{ backgroundColor: '#f8fafc', color: '#0f172a', borderColor: '#cbd5e1' }}
                            />
                            <div data-pdf-hide="true" className="text-[10px] text-slate-500 font-bold flex flex-wrap justify-between items-center gap-2 print:hidden">
                              <span>* O texto acima ser√° impresso no relat√≥rio oficial em PDF.</span>
                              <div className="flex items-center gap-3">
                                <button
                                  type="button"
                                  onClick={() => setQuantitativoCriticalAnalysis('')}
                                  className="text-slate-700 hover:underline flex items-center gap-1 font-bold"
                                >
                                  <RotateCcw size={10} /> Recalcular pelo Gr√°fico
                                </button>
                                <button
                                  type="button"
                                  onClick={() => setIsEditingQuantitativoAnalysis(false)}
                                  className="text-slate-900 underline font-black hover:text-slate-950"
                                >
                                  Concluir Edi√ß√£o
                                </button>
                              </div>
                            </div>
                          </div>
                        ) : (
                          <div
                            onClick={() => setIsEditingQuantitativoAnalysis(true)}
                            title="Clique para editar o texto da An√°lise Cr√≠tica"
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
                          CEO - Centro de Especialidades Odontol√≥gicas.
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
                            Personaliza√ß√£o Institucional
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
                          Anexo de Papel Timbrado dos Relat√≥rios
                        </h3>
                        <p className="text-xs sm:text-sm text-slate-500 font-medium mt-1 max-w-2xl">
                          Anexe a imagem oficial do papel timbrado do √≥rg√£o ou institui√ß√£o (contendo cabe√ßalho, logomarcas e rodap√©). A imagem anexada ser√° inserida automaticamente no topo de <strong>todos os relat√≥rios exportados em PDF</strong> (Estoque, Cat√°logo, Solicita√ß√µes, PCA, Termos de Doa√ß√£o e Recibos).
                        </p>
                      </div>

                      {letterheadImage && (
                        <button
                          onClick={() => handleExportInventoryPDF()}
                          className="px-4 py-2.5 rounded-xl bg-slate-900 text-white font-extrabold text-xs flex items-center gap-2 hover:bg-slate-800 transition-all shadow-sm shrink-0"
                        >
                          <Download size={15} /> Testar Exporta√ß√£o PDF
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
                                <p className="text-[10px] text-slate-500 font-medium">Sincronizado e pronto para emiss√£o</p>
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
                            <li>Utilize imagens em alta resolu√ß√£o com fundo branco ou transparente.</li>
                            <li>O timbrado √© posicionado no cabe√ßalho superior de cada p√°gina gerada.</li>
                            <li>Sua altera√ß√£o √© salva imediatamente para todos os administradores.</li>
                          </ul>
                        </div>
                      </div>

                      {/* Live A4 Preview Simulation */}
                      <div className="lg:col-span-7 space-y-3">
                        <div className="flex items-center justify-between">
                          <label className="text-xs font-extrabold text-slate-700 uppercase tracking-wider">
                            Pr√©-visualiza√ß√£o da Folha A4 com Timbrado
                          </label>
                          <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest bg-slate-100 px-2.5 py-0.5 rounded-md">
                            Propor√ß√£o A4
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
                                    Cabe√ßalho do Timbrado Oficial
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
                              <span>Relat√≥rio Oficial do Sistema</span>
                              <span>P√°gina 1 de 1</span>
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
                <h3 className="text-2xl font-black">Gerenciamento de Usu√°rios</h3>
                <button 
                  onClick={() => setIsRegistering(true)}
                  className="bg-[#1C1917] text-white px-6 py-3 rounded-2xl font-bold flex items-center gap-2 hover:bg-[#292524] transition-all shadow-lg"
                >
                  <Plus size={20} /> Novo Usu√°rio
                </button>
              </div>

              {isRegistering && (
                <div className="bg-white p-8 rounded-[32px] border border-[#E7E5E4] shadow-sm max-w-2xl">
                  <div className="flex justify-between items-center mb-6">
                    <h4 className="text-lg font-bold">Cadastrar Novo Usu√°rio</h4>
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
                          placeholder="Nome do funcion√°rio"
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
                        <><Save size={20} /> Salvar Usu√°rio</>
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
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">A√ß√µes</th>
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
                  <h3 className="font-bold text-[#1C1917]">Itens Exclu√≠dos</h3>
                </div>
                <div className="bg-white rounded-3xl border border-[#E7E5E4] shadow-sm overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse min-w-[600px]">
                    <thead>
                      <tr className="bg-[#FAFAF9] border-bottom border-[#E7E5E4]">
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Item</th>
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Exclu√≠do em</th>
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Exclu√≠do por</th>
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">A√ß√µes</th>
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
                  <h3 className="font-bold text-[#1C1917]">Solicita√ß√µes Exclu√≠das</h3>
                </div>
                <div className="bg-white rounded-3xl border border-[#E7E5E4] shadow-sm overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse min-w-[600px]">
                    <thead>
                      <tr className="bg-[#FAFAF9] border-bottom border-[#E7E5E4]">
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Solicita√ß√£o</th>
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Exclu√≠do em</th>
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Exclu√≠do por</th>
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">A√ß√µes</th>
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
                                if (window.confirm('Deseja restaurar esta solicita√ß√£o?')) {
                                  await updateDoc(doc(db, 'requests', req.id), { 
                                    deletedAt: deleteField(),
                                    deletedBy: deleteField()
                                  });
                                  setToast({ show: true, message: 'Solicita√ß√£o restaurada!', type: 'success' });
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
                      <p className="text-[#A8A29E] text-sm">Nenhuma solicita√ß√£o na lixeira.</p>
                    </div>
                  )}
                </div>
              </div>

              {/* Deleted Transactions */}
              <div className="space-y-4">
                <div className="flex items-center gap-2 px-2">
                  <History className="text-[#78716C]" size={20} />
                  <h3 className="font-bold text-[#1C1917]">Movimenta√ß√µes Exclu√≠das</h3>
                </div>
                <div className="bg-white rounded-3xl border border-[#E7E5E4] shadow-sm overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse min-w-[500px]">
                    <thead>
                      <tr className="bg-[#FAFAF9] border-bottom border-[#E7E5E4]">
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Movimenta√ß√£o</th>
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Exclu√≠do em</th>
                        <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">A√ß√µes</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E7E5E4]">
                      {transactions.filter(t => t.deletedAt).map(trans => (
                        <tr key={trans.id} className="hover:bg-[#FAFAF9] transition-all">
                          <td className="px-6 py-4">
                            <p className="font-bold text-sm">{trans.item_name}</p>
                            <p className="text-xs text-[#A8A29E]">{trans.type === 'entry' ? 'Entrada' : 'Sa√≠da'} - {trans.quantity} un.</p>
                          </td>
                          <td className="px-6 py-4 text-sm text-[#57534E]">
                            {trans.deletedAt && new Date(trans.deletedAt).toLocaleString('pt-BR')}
                          </td>
                          <td className="px-6 py-4 text-right">
                            <button 
                              onClick={async () => {
                                if (window.confirm('Deseja restaurar esta movimenta√ß√£o?')) {
                                  await updateDoc(doc(db, 'transactions', trans.id), { 
                                    deletedAt: deleteField()
                                  });
                                  setToast({ show: true, message: 'Movimenta√ß√£o restaurada!', type: 'success' });
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
                      <p className="text-[#A8A29E] text-sm">Nenhuma movimenta√ß√£o na lixeira.</p>
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
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">N¬∫ / Data</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Setor</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Status</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Itens</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">A√ß√µes</th>
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
                            {req.status === 'EM_SEPARACAO' ? 'EM SEPARA√á√ÉO' : 
                             req.status === 'DEVOLUCAO_PENDENTE' ? 'DEVOLU√á√ÉO PENDENTE' :
                             req.status === 'DEVOLUCAO_APROVADA' ? 'DEVOLU√á√ÉO APROVADA' :
                             req.status === 'DEVOLUCAO_RECUSADA' ? 'DEVOLU√á√ÉO RECUSADA' :
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
                    <p className="text-[#78716C]">Nenhuma solicita√ß√£o encontrada.</p>
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
                  <h3 className="text-lg font-black">Solicita√ß√µes de Devolu√ß√£o pendentes de aprova√ß√£o</h3>
                  <p className="text-xs text-[#78716C]">Visualize e aprove o retorno de materiais ao estoque.</p>
                </div>
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-[#FAFAF9] border-b border-[#E7E5E4]">
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">N¬∫ / Data</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Setor</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Status</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Motivo</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Itens</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">A√ß√µes</th>
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
                            {req.returnReason || 'N√£o especificado'}
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
                    <p className="text-[#78716C] font-bold">Nenhuma devolu√ß√£o encontrada.</p>
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
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">N¬∫ / Data</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Status</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider">Itens</th>
                      <th className="px-6 py-4 font-bold text-sm text-[#78716C] uppercase tracking-wider text-right">A√ß√µes</th>
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
                            {req.status === 'EM_SEPARACAO' ? 'EM SEPARA√á√ÉO' : 
                             req.status === 'DEVOLUCAO_PENDENTE' ? 'DEVOLU√á√ÉO PENDENTE' :
                             req.status === 'DEVOLUCAO_APROVADA' ? 'DEVOLU√á√ÉO APROVADA' :
                             req.status === 'DEVOLUCAO_RECUSADA' ? 'DEVOLU√á√ÉO RECUSADA' :
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
                                  title="Editar Solicita√ß√£o"
                                >
                                  <Edit2 size={18} />
                                </button>
                                <button 
                                  onClick={() => handleDeleteRequest(req.id)}
                                  className="p-2 text-rose-400 hover:bg-rose-50 hover:text-rose-600 rounded-xl transition-all"
                                  title="Excluir Solicita√ß√£o"
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
                    <p className="text-[#78716C]">Voc√™ ainda n√£o fez nenhuma solicita√ß√£o.</p>
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
                    {editingRequest ? 'Editar Solicita√ß√£o' : 'Nova Solicita√ß√£o'}
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
                      Cancelar Edi√ß√£o
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

                            return filtered.map(group => {
                              const groupBatches = items.filter(i => !i.deletedAt && normalizeString(i.name) === normalizeString(group.name) && (i.quantity || 0) > 0);
                              const groupTotal = groupBatches.reduce((s, i) => s + (i.quantity || 0), 0);
                              const hasMulti = groupBatches.length > 1;

                              return (
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
                                    <div className="flex items-center gap-2">
                                      <p className="font-bold text-[#1C1917]">{group.name}</p>
                                      {hasMulti && (
                                        <span className="px-1.5 py-0.5 bg-amber-100 text-amber-800 text-[9px] font-black uppercase rounded">
                                          {groupBatches.length} lotes
                                        </span>
                                      )}
                                    </div>
                                    <p className="text-[10px] text-[#A8A29E] uppercase font-black tracking-widest">
                                      {group.category} ‚Ä¢ Saldo: {groupTotal} un
                                    </p>
                                  </div>
                                  <div className="flex items-center gap-2 text-emerald-600">
                                    <Plus size={16} />
                                    <span className="text-xs font-bold">Adicionar</span>
                                  </div>
                                </button>
                              );
                            });
                          })()}
                        </motion.div>
                      )}
                    </div>
                  </div>

                  {requestBasket.length > 0 && (
                    <div className="space-y-3">
                      <label className="block text-xs font-bold text-[#A8A29E] uppercase tracking-widest">Itens na Cesta</label>
                      {requestBasket.map(item => {
                        const productBatches = items.filter(i => !i.deletedAt && normalizeString(i.name) === normalizeString(item.product_name) && (i.quantity || 0) > 0);
                        const hasMultipleBatches = productBatches.length > 1;

                        return (
                          <div key={item.product_id} className="p-4 bg-[#FAFAF9] rounded-2xl border border-[#E7E5E4] space-y-3">
                            <div className="flex items-center justify-between">
                              <div>
                                <p className="font-bold text-sm text-[#1C1917]">{item.product_name}</p>
                                {hasMultipleBatches ? (
                                  <span className="inline-flex items-center gap-1 text-[10px] font-black uppercase text-amber-700 bg-amber-100 px-2 py-0.5 rounded-md mt-1">
                                    üè∑Ô∏è {productBatches.length} Lotes Dispon√≠veis
                                  </span>
                                ) : productBatches.length === 1 && productBatches[0].batch_number ? (
                                  <span className="text-[11px] font-medium text-[#78716C]">
                                    Lote: {productBatches[0].batch_number} (Venc: {productBatches[0].expiry_date || 'Indet.'})
                                  </span>
                                ) : null}
                              </div>
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
                            
                            {hasMultipleBatches && (
                              <div className="pt-2 border-t border-[#E7E5E4] flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                                <label className="text-[11px] font-bold text-[#57534E]">Escolha o Lote Desejado:</label>
                                <select
                                  value={item.batch_id || ''}
                                  onChange={(e) => setRequestBasket(requestBasket.map(bi => bi.product_id === item.product_id ? { ...bi, batch_id: e.target.value } : bi))}
                                  className="text-xs font-bold bg-white border border-amber-200 text-amber-950 rounded-xl px-3 py-1.5 focus:ring-2 focus:ring-amber-500/20"
                                >
                                  <option value="">Autom√°tico (FEFO - Primeiro que vence)</option>
                                  {productBatches.map(b => (
                                    <option key={b.id} value={b.id}>
                                      Lote {b.batch_number || 'Sem Lote'} ({b.quantity} un - Venc: {b.expiry_date || 'Indet.'})
                                    </option>
                                  ))}
                                </select>
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  )}

                  <div>
                    <label className="block text-xs font-bold text-[#A8A29E] uppercase tracking-widest mb-2">Observa√ß√£o (Opcional)</label>
                    <textarea 
                      value={requestObservation}
                      onChange={(e) => setRequestObservation(e.target.value)}
                      placeholder="Alguma observa√ß√£o importante?"
                      className="w-full px-4 py-3 bg-white border border-[#E7E5E4] rounded-2xl focus:outline-none focus:ring-2 focus:ring-[#1C1917]/10 font-medium min-h-[100px]"
                    />
                  </div>

                  <button 
                    onClick={handleSubmitRequest}
                    disabled={isSubmittingRequest || requestBasket.length === 0}
                    className="w-full py-4 bg-[#1C1917] text-white rounded-2xl font-bold hover:bg-[#292524] transition-all shadow-lg disabled:opacity-50 disabled:cursor-not-allowed flex items-center justify-center gap-2"
                  >
                    {isSubmittingRequest ? 'Enviando...' : <><Save size={20} /> {editingRequest ? 'Salvar Altera√ß√µes' : 'Enviar Solicita√ß√£o'}</>}
                  </button>
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'devolution' && (
            <motion.div 
              key="devolution"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="max-w-5xl mx-auto space-y-6"
            >
              {/* Refined Banner Header */}
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-5 bg-gradient-to-r from-amber-500/10 via-orange-500/5 to-transparent p-6 sm:p-8 rounded-3xl border border-amber-200/80 shadow-xs relative overflow-hidden">
                <div className="space-y-1.5 z-10">
                  <div className="flex items-center gap-2">
                    <span className="p-2 rounded-xl bg-amber-600 text-white shadow-md shadow-amber-500/20">
                      <RotateCcw size={20} />
                    </span>
                    <h2 className="text-2xl font-black text-slate-900 tracking-tight">
                      Devolu√ß√£o de Materiais
                    </h2>
                  </div>
                  <p className="text-xs sm:text-sm text-slate-600 font-medium max-w-xl">
                    Gerencie o retorno de materiais do setor <strong className="text-amber-700">{selectedSector}</strong> ao almoxarifado de forma simples e organizada.
                  </p>
                </div>
                <button
                  onClick={() => {
                    setDevolutionBasket([]);
                    setSelectedDevProduct('');
                    setDevolutionReason('N√£o teve uso');
                    setDevolutionObservation('');
                    setShowDevolutionModal({ show: true });
                  }}
                  className="z-10 bg-gradient-to-r from-amber-600 to-amber-700 text-white px-6 py-3.5 rounded-2xl text-xs font-black uppercase tracking-wider hover:from-amber-700 hover:to-amber-800 transition-all shadow-md shadow-amber-600/20 flex items-center justify-center gap-2 whitespace-nowrap self-start md:self-auto hover:-translate-y-0.5 active:translate-y-0"
                >
                  <RotateCcw size={16} />
                  Solicitar Devolu√ß√£o
                </button>
              </div>

              {/* Statistics KPI Row */}
              {(() => {
                const devRequests = requests.filter(r => r.sector === selectedSector && r.isReturn && !r.deletedAt);
                const pendingCount = devRequests.filter(r => r.status === 'DEVOLUCAO_PENDENTE').length;
                const approvedCount = devRequests.filter(r => r.status === 'DEVOLUCAO_APROVADA').length;
                const rejectedCount = devRequests.filter(r => r.status === 'DEVOLUCAO_RECUSADA').length;

                return (
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 sm:gap-4">
                    <div className="bg-white border border-slate-100/80 rounded-2xl p-4 shadow-xs flex items-center gap-3 hover:shadow-md transition-all">
                      <div className="p-3 bg-amber-50 text-amber-600 rounded-xl">
                        <RotateCcw size={18} />
                      </div>
                      <div>
                        <p className="text-[10px] font-black text-slate-400 uppercase tracking-wider">Total Enviadas</p>
                        <p className="text-xl font-black text-slate-900">{devRequests.length}</p>
                      </div>
                    </div>
                    <div className="bg-white border border-slate-100/80 rounded-2xl p-4 shadow-xs flex items-center gap-3 hover:shadow-md transition-all">
                      <div className="p-3 bg-yellow-50 text-yellow-600 rounded-xl">
                        <Clock size={18} />
                      </div>
                      <div>
                        <p className="text-[10px] font-black text-slate-400 uppercase tracking-wider">Pendentes</p>
                        <p className="text-xl font-black text-yellow-600">{pendingCount}</p>
                      </div>
                    </div>
                    <div className="bg-white border border-slate-100/80 rounded-2xl p-4 shadow-xs flex items-center gap-3 hover:shadow-md transition-all">
                      <div className="p-3 bg-emerald-50 text-emerald-600 rounded-xl">
                        <CheckCircle size={18} />
                      </div>
                      <div>
                        <p className="text-[10px] font-black text-slate-400 uppercase tracking-wider">Aprovadas</p>
                        <p className="text-xl font-black text-emerald-600">{approvedCount}</p>
                      </div>
                    </div>
                    <div className="bg-white border border-slate-100/80 rounded-2xl p-4 shadow-xs flex items-center gap-3 hover:shadow-md transition-all">
                      <div className="p-3 bg-rose-50 text-rose-600 rounded-xl">
                        <X size={18} />
                      </div>
                      <div>
                        <p className="text-[10px] font-black text-slate-400 uppercase tracking-wider">Recusadas</p>
                        <p className="text-xl font-black text-rose-600">{rejectedCount}</p>
                      </div>
                    </div>
                  </div>
                );
              })()}

              {/* Modern Segmented Controller */}
              <div className="bg-slate-100/80 p-1.5 rounded-2xl inline-flex flex-wrap gap-1 w-full sm:w-auto border border-slate-200/60">
                <button
                  onClick={() => setDevolutionSubTab('my_returns')}
                  className={`flex-1 sm:flex-none px-5 py-2.5 rounded-xl text-xs font-black transition-all ${
                    devolutionSubTab === 'my_returns'
                      ? 'bg-white text-amber-700 shadow-sm border border-amber-100/80'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Minhas Devolu√ß√µes
                </button>
                <button
                  onClick={() => setDevolutionSubTab('sector_stock')}
                  className={`flex-1 sm:flex-none px-5 py-2.5 rounded-xl text-xs font-black transition-all ${
                    devolutionSubTab === 'sector_stock'
                      ? 'bg-white text-amber-700 shadow-sm border border-amber-100/80'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Estoque do Setor ({selectedSector})
                </button>
                <button
                  onClick={() => setDevolutionSubTab('eligible_deliveries')}
                  className={`flex-1 sm:flex-none px-5 py-2.5 rounded-xl text-xs font-black transition-all ${
                    devolutionSubTab === 'eligible_deliveries'
                      ? 'bg-white text-amber-700 shadow-sm border border-amber-100/80'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Entregas Eleg√≠veis para Devolu√ß√£o
                </button>
              </div>

              {/* Subtab Content */}
              {devolutionSubTab === 'my_returns' && (
                <div className="bg-white rounded-3xl border border-slate-200/80 shadow-xs overflow-hidden">
                  <div className="p-6 border-b border-slate-100">
                    <h3 className="text-base font-black text-slate-900">Hist√≥rico de Solicita√ß√µes de Devolu√ß√£o</h3>
                    <p className="text-xs text-slate-500 font-medium">Acompanhe o andamento e o parecer das solicita√ß√µes de devolu√ß√£o do seu setor.</p>
                  </div>

                  xúÏ}Ms€Hñ‡}Eñ ;§∫Mä¢$hdyUíÏvèm©$π¶7"R⁄ ¿@}4K±á=Ìiˆ¥∑⁄>tÃDÙ©c/s’?ô_≤Ôe&ÄêôHPî-ªåÓíI»œ˜^æÔ∑·zgd‡;q¸⁄—'„Œ*âGÎ„ŒèùÌ\vV6ˇQ\”à˛iB„$V˛JH˜ƒÛµ#ÚdìD›òí>?yBbÍ√ÍÚ[ˇ≥–d¯Ìª®Î¬#ƒV≤®k}‰å€0l|™yÜêAƒ	Å«^$tì'ƒÒ˝>lv'£«Èu≈úﬁ{.)|ÌzÓ‚?j;à¯†€⁄Ÿ¿%6¸N»Gz˘d ª∫2?)Ì‘Ò∞˚NB;kΩ•=rF.çƒ?‚ó~è˝tû—h]¸‚åé·/¸B¢p∏‘ÌÙ/|ím˚I"'àΩƒÉ,V˙ÚPºù-˜»âO/ÿüŒ ÙÒEˆ9
œÒ≥á€– ñ¸q'ﬁ…eÁò&ÁîdË@WÜ9™a≠∞ñ“"§0⁄Á£Y÷Ä™æÅ|&Áë3&Ö±„X˚µ-Bõ0ä†–h$ùcﬂ|$	ΩH:1ˇóÔ cX{XœÛSËåå/:˝Ó_vñ≥Òá∫›¥!ﬂPÍ∆æ7†ÌŒ⁄b7	ﬂå«4⁄vb⁄^º™ü–ŒhÜâ+fª≥eÎ”ëw˙Æ’˛Û_˛/ôÙúÏ@àÊ]>‡D^Ü«ßx˚0âº`ÿnçìŒ≠[ú÷Ùõœ€Âﬁ¯‚ë66o•∏u'@ô	Æı ÷qiŸ9˜ÚÌ$˜ÙÙ*øpŒq‚$ìò¢÷ŒÓO{/ﬂloÌΩﬂﬂ}Ω≥˚˙h∑EûíñÑò|’˘∑áΩî"‰ﬂ"Î7Íwkˇ`Ôß≠ù≠¥_:¢ë„ªYœÈw©ÔÙ÷z?ÿ›~s(ıÖ1Õ∫f_§~ŸwÎN[)]ÓıdÿïZÃP∞U€‡’á+üZÓp˛≈f.∂˚'æÏ›∏’¬æà/Õ[ùÚ≤eÂßÚu‚ê1¶„9Î@IÃ
H‘KF≤GúêÛ/ :Kèz=Ô¡?)YŸQ=B^ÖâwÆWgQø@ˆKTK*7ñ‡î‹¸/7;JÒÙ\∆•H,éeæoú%C∆œa‰»Ï7éÒP¯c¢gRæÖ∫37=ì5áØLpÎ„S«œÅ}ä-7óoÖÓdêº`tW0¯$
Éae¿_ÿlÛ˜˛4qÇƒK.ﬂï∫Wã∞„Ïuõ˛≠°cN‡¡@8<éitÊ 7iáác~¬K`m≤C8µl/ÅüUm·≤%ˇ¥P˘ïâ_Mód\◊tÕ≤⁄,Í∆Ò$I`%kz
ÉmXûèO¶ÌEƒ°ò&áß·π{vh‚x˛´–u¸ˆ`8<_6eBÔWÎ¯Å\’BÅJyúbI∆‹Æ"öı%RõS≈»L»»±´$ÜdÉ∑@x<È–¿EQÉ}‹:ÎíÀAà»ºku€ı4âkÂüRùÄÀØç%æ+¶≈ÊÍ~÷ ó∏fŸ{NBv◊ß¡09e/ˆLòZ&˙@„˚|C≈6îP5–®∑q@∑ÁrÀ£ãé3IBπ¡¿`{ñ|µwEñMé´ÚXvz«£ì∑∞˘ößìëC‚0»Àı_Øˇó¬ˇœB¬ø“‡Ãs\ßkDv∂∞˘&Ò|ò	ÅD%ÿ∂3†˚ÖC—?¬^÷Ò;ëh9ÙÄ‹√o!™¬(`#t¸·dDF–t‰9>q†-^8ëw(c¨*ïÿØy^yª
∏Sæîà—áì„#Áòsíbﬂ«I8¯ÿRC^Ê≤É:•)+äﬂÁ<YN;êÃú¯˘‘s](a≤
·“Vèmîç”ï
£@X÷ddsaòü &¿ˆÏ¬Z rÿﬁC‹f“ûqœ˝”Mœ6Á'≈à∫ﬁd¥∞yàm√¶PÜG90'N‰¬‚ë <Ip8^0'±wF…û∆ãBŸ√€°˛Ã´nÇΩZ≤Õè=ùêÕUÇÉI¡‚zà¿GûpÌOJQôFPãÊﬂy9ıD®’>ÿˆ∫~8‡êÇ,ˇ¸3ión1¥xÊD£Î_û√p¢‘ÜÙÎ‚¢±w/c…&ÈiÉ„GÛãwB⁄’ï*ãÌ´≠rtŒÁ	kwpÃ“πù(¨—ôNVàÿ4GlÜIÊÉ§⁄°Ó0˘∂p‘!')X,qLáq<Ò";>¸ûè‡>	c–óhƒŒ†?M<	ë≈©Á!°píöcÔ<é¯Ù(‡OñÎtˇÙbÏE‘EÏéw˘gˆÆAßœ_®±.ŸªØ≥ØÈÎü¿&ê…≥ñFÅÈq⁄Èˆo®∆ØWd¶ã_+?tzKkE^—Ü¡ÓÆ†zØ∂·uy-áëY:V´JTï1≈v$-•…∆∆VcV:^}0…j≥TPssóM*EjÀ(é-L≠b≠,dØ¿…Já!¿«-h
Ûì´`Ú©*odÚ)œ‰ìj°ÛœüŒk˝ÚçŒU˚Ç>,˜ ∆ŒîhîØfõÉ÷
å⁄h,ŸT∂|%G¿,Cü
Nbπèú˘â3«Û‹ä"9ªÖ’ñ¨Gíbª¥ﬁ’ªkÓ¯H+îÏG◊ø yzŒK?˜’.XÃrÉZu≈”_“5o¥‚ñ+ª;"¿Á›A3Ccâ∂vd?2ôàÅ
‚Bﬁ^◊:%Z.ù5©lu%»Ô$í˜#Íƒìà¢d◊Çn´·itÏ$É”˜¡ëÃ˙D⁄dÛóab=çá˘4‰”¡n‘àBxŸú¨v<\æG>ìfÀ˜æCôˆ@2PÉë8Æ£—igû¢~uˆ”˘ÙÔIU+?]0C±∏ì≠NÌ∏πBy–f_ÑyÆbΩîxk¶ÉzG4w2ΩﬁN¸ë&Ì∑6˛©Ãs◊âsÓ7xŸ<Ò"~¥y5≈VÒZ˙’Ê’ësÒ„L/¶™ù≈^‰S≠w"xgPÛ÷•≈Áv‚∂·Ça∞˝E˘ÑûQ2â√V”∂˜rcò‘¡áW©ûPËQ≈poVî˘Ä√¥⁄≥Ú—UçY‰™∂ù+{˚÷ÙÉï!ãüÎ76[È$Ë⁄¬q»Z‚.»‹r¡vòs‹π%œBºM\ŒÃôJ©≈∫Ì1 ¥ıÊ;…¶$xƒU∆#ÓH*Í-IE˝M|Í_Æ’ûz∑bÑ°æ7Ùé}˙ﬁÖO∞:ç’∂ò› âË–â	Uﬂ˚30$13«‹éÌçèÿM{F≠Ï±«:>vºãê+fcÒÇ	˝-Õå.ü÷»2WﬂoŸ≥l˜ı—¡ÓÛ7ª≠oé‡¶'?Ω#¯WÊÛ˝5xz2ııx|3Go~1√™ï‚Qæ¿˜——∫ÃÅñ⁄´c@≠∫['∂#ÙQvVœÊŒk3Óéì¶ç∏«ÈØ„G‘q/π7 ï\≈?ˇLzıí@N÷Gé¿2Wt∆ 0ûAÉùrﬂ6≠[–Ù¸‚®aıh#Î^~U"¨ºZƒ+_Õt≥~/ï(îÆÙ´’˜⁄INap√”™_}ìé∏®¡Èj…+∑|∫ILgˇ"Æ˜ﬁÁóö‚eßZUª˛∂ßŸJ_-M’p{E\/wuºOÒ≤%‚i$øÍ@k$Èëœ¢ô‚§·ò©§R¶Ø@æ8´w‚ys%]rs7#\ºmnL˜D„äì†ﬁﬂ1»+83ej5Œ˙z¿µÄùår⁄Y(ufß,´K£µk"Wf˚a˜b™l¯ZE-£ÊZxèPqÁﬁˆﬁ=ÌÚZ6Ï ªÀ`P†?Ë,fÚpI/Ùì0Ã⁄1LZâ”¸(t‚§Ωp∫aú{&π }9©<L˛x˝úuë3"Iò8˛˘v!˙¢bw·>YÇìp°xZ,S˝®‡“íÃ®≈ù£æv^:’JÄ¿5¨≈a”&~©¥çJ]l<RË\5⁄ŸõÖd∫F°'˜æû@;çÕ∑®Uê1Xò·ıﬂŒ®w‡à¯úO4¡!ùﬂÏqD^Ï\≥Lï]¸…ıﬂPôÿ^H∑kaë9¢z1ê)Gˆﬂæ˛˜éœıù¬]Ù3∆üÖàÃ›“£“cK[ÅáﬁÎ˚çaçi˙ºÁs∆âNó~CÖã…oñ“◊ßHÎ∂\ó˝PÑ÷äÌ] kÊ∞Ë¶‚a4à˘¬øné·Œ$Bò¯3)W5N©éL|w@y#üjë1&#Òˇ…H3™÷IØ˚¯>	!ˇ“+ì]á/àÙ¬≤Ù¯r˘qïjˇú˚ó Ø”9Ô¨˘M)Ò€UÊ÷ÿ*…X¸¡”Œ€«Ω≥”wπöˇí!ûLQã˚≠‘TîUâÖ5wU¿ÜæF ¥Tîıwoø_ﬁ^~º¸∑†ñ«•9˝VÎ˛H˙ˆ˚áè.?ÿ~W‘˜o£%∞íåÆˇ√Oº±/q;ÃpFˇ¨@6%vhÂ&u¯_
ÿÌÄz™¿Õ˘ÌÁÁÈ€Ôü≠¡ˇVﬂΩç§√u˙aóœJ≈:˝AêŸ˛j5B˜Ì˜[è∂˙èwﬂ-T)∞˙dT	ó¿#é3nçº‰…Ù‘	\ü¬‹ëªRië´Ñƒa;ç`müy‘wcÇ'∆^‰¥"ÎØjÉı⁄Çˇ=Fºxî£ JQ•®Ì}˚˝Ó√›µ]X⁄a¸=˛¡≈å;Àd‰ÆÁ_˚ƒJ_Wô⁄ÓÅ‚Ï+^É7:®–Ë˜6|ÁŒ#y
~(i≥´í∂∆›‰N¿K8òüÖ ≠®ã3÷â≤{/O4z^‰L–ñ≠¸—ás6·{®Föá–5lR˘ï\é§-®{√È)úÊ4z≤:Ÿ◊Q‚sP∑*-õ í)√∫¢'œ˜\‚gO¬¡$^GE1Z5Ú/9bIwR~C=ñ3«ü nO¸èH∫.ªÈ‚®Ÿp†ÄCxÉ
*ÒC˙f{⁄Ìv≥vÓì¥°uBªâi“eùïå Û^¡MiXÄ;ÃGﬁò["!ÌmÓfÌ9ã&–Vùg&C„?^”s—˙ÂÜz'_U@9j˚aÔé_¬û™vˇ˘Ä_%<>s ñMœ¥	Úu«ÈòD◊*Êº.æ‹.¢âQëä<—3úØˆì<RØ:Â€»ü3-C#ı*™}J÷M"o‘^¨W˙¿r•ËC„6»+gÿﬂV9ó›ì(a√Ëh—~á?ﬂ'ÍÆﬁ-.÷™ LT-çïX◊¥_Ø‡PÏºÖ≤EceD|ô˘m$’-eE`UA’(õc§VåÍâ}ßWèÃ≤tç∂√$X|Ò7=rEÜ	[Ó9´úÍ6!>ŒˆäDÀd˛1A¿tê!73≤¿◊˙<>·òÒ◊Ã*
o\•”∆œõÏÔ∆∆‘∑)sÕ∆ﬂ˝œALuèZU„·†∆ãbj	kµ∞YïÊÓ ÇœMmtßY«Ω»“ëâQ4—•;-a ö=Mj≤4à7S¶@uê¡•öi0·]‡*÷A≤∞πÕ>%°ôäî^áméúÖÕ}¥≈%!Ÿ≈Øçp√Äô`6wBûQ∆Ù∫ûB©’?J–G›7vøπÅ‚"Uá¨Zhj5sj;]≠(}‚ë^ßG¨ÖMúú#‰± µr´ MΩûñgT‹q]Ñ]\◊É\Ó∆ÑsÜdôxD7rÈ)π®q_Ãiº¸p—¶ö®îeBœ]Á∑\O∏ÔMí(dp¶Ñc›∑•÷ô ò€j‘–î8«>Uê«‘Øà¢E#£çH˙Azõ—FrJW{Z%Qeè}zíòdÏ‰¥†TÕvŸö©AMÁätŒ°ÅG¨Öëª.nÙô:~a3’^·m,%ß˙!*¯Jn”{Ea≥¥≤á∫HÄO0—ÛŒ Él∂À|r©&&ç=2O–ƒ~íQ˜˘®ﬂ<‹tâÏééªÊ1œ4∞Á—ÆrxER±Ù0˜O2#{J–L‹.&ë[~ûPµ~Ωty‚U„–”|ZØÆˇ‹¬–Ì@¢+ÜÁ~∂¨ˆ6”(Ÿœ∑
èêü¢◊…õ‡6–¿r∞¶Œ·∑HsH,Nâç‰8t/uçJ¥÷À|ô˚’}‚Ë^,öe<Å¥Iná ˛c£˜F‚ññ≥_{êòÖˆZ2^F#N~qsŒ¨ﬁîì_v:jºî÷ù1ÁÏÎﬁUä`+ˆQ¡jœ%y`[â^E◊`üxÙàïé/!õÂ)pjû/âgì1F.§\n;»&-l´uü4–h5÷x°åÚ„8kÇ˘»)ıSë]é»à±lÔ2K©ôπ.xï˘œQí•? ûX§)Ÿw˚˙Hû\
Á~o˚è,êJ5U»FΩ∏UiUùò‰ëöä÷(7`ß"T§∂&^∑w=üæmçÜ Q≠—Ä˝ÛÊ˛Âw|¸˚_[Ôid¿aô<ìı¨û·ÿÅùóøΩÊOæ{¨ßóO≥¸îØôvÓê-l6>Æ¸í28b[vNπÚÀﬂIowi‡∆ˇÏ%ßÌiYÿzäó<•ﬂûíñ˝4Ã¶èöûp£õtUGÂı¥o◊vµn≥ÈU<ÕóÛ<LÉ√NêæÖü+V¯-…Ù#:ÛxPAó≤öùj3aá∂Ò2ømÄù6~≥ÈUü‹ﬁ*ÆÖ=ˆÂêÈù0vÜt‘åJ£ü-ß‘ a>ríΩñ\ÓuÂ˚ÈN«›0¶∑J«±ÉotºÓÂØÄé„Fˇ Ë∏§6Æ®wÛ+eR.?q'àπ=ä~bnÒê±´ç•D´8∆k^¢äIÜ_∂í.,| ¯e)∑JÈÊà–;´Ñ\îësIÛ=û	<∂Æúl≈ÊRÛM%h‹*•d2›n◊∆]BŸÃ˛ﬁ¡—÷¡ã-≤≤∫∫∞)õπ…ÌΩW˚/^ΩÿŸ[ÿÃ?œ‹‹÷´˝Ωó[õ¸ﬂôõ9‹{˘Ê˙^ˇTˆÒÊçë√˝É≠ˇ.5…o‹`;^mÌl·F‡ø37Û|Ô¡ø7ÿ≈ó◊ˇÎ‡€C˛…∂©z◊˛îôÜö	∞Å0ZÈ2Ì»¢%Q¸íÜµ`∂ï⁄Æ⁄o^/"ÉøÔB‡4⁄˚€GÏ˚∂„]8§Ω˝ˆÌY‰ƒÉê¥ü∞Ø[£qË√ØÄ≠Ï˚^8C∏Ò√·svc78£~∏Ω˚˙'vÁπ„ck˚˘Kˆı Ù·À¡K—wÛ¶_Q¥B∑_±/?N<|Ïüx£/=ˆìhÄ∆#ÑV∂ƒ•q[ù´µeìd·)—º]X$õñ÷Iãô—k·5S‰⁄ÊEÄÌy§x(’Û¿(H∞7ëy·√∑ì¥•ºê∏Y6¡ WÑ˙ SŒ≥_òîM«u{u√úÖEä+ÉTnA∆o÷î∑tòHàGäÙm∆sÃ«–∂ÙÛåçIÑcaS˙2cs2·fA˙6cÉ µ∞Y¯:cìE⁄á!áÚ˜ï»'˘˘óõÀ»Ô¬fˆqf¯„ƒaèö±°åÚ/lfgl*?7∞éN˙y∆∆≤c=‘ƒ«Yó-Xr¸g∆&}^ÿ‰ŒZm◊b]9ê,ád«∂Mø˚1V∆D+≥x€u…zΩèAÃ;Ò–•s–;>*wk¬ú¯UÂL—÷YH¶´ÇÀ8Ã)[$0™u<®'€LnØÓ ¸ÑˆO–PbØ±zﬂHÅtS≈œ\}58‰ÚruõÉ|≤∞‹»;„«§Vv©∞$•©Äﬁæ™êÜÖw|ﬁRöÃnT‘∞>^Î…fô<≈åÏÅfÈ∫3z-†ë'yxü¶cZƒÃzò∆]˘Îº\= ÌVPÄ≥˝XËF
£;QL_I9ﬁÒ&Ó!7¡Ñﬁù«ÑZbW¿tXº!*§•Ã¥|,ê m£Ñ¸vÜ ,Z"œÜ3+Ë√2i¿^˛e^ /∑˘≈Ä;w&ùîœË>án¨_ùœ\°Œú†KnsÆ>t≥AÕj]RiΩGCÊ Wo`jƒ˛‚¬’ÛÆc∞áõ∫æ∆ÔΩºVPBﬂÛ#ı|†(óπµı”´≤9≠ãTL@åõ1¬R]ï˘q¿R£sgÄU1ñöR'hôè√®3=K«[∞KopJè√CwlV»iqõÇ…0r.ŸiZ !e”8‰Ê†›ÄÀSw!o∑ìÖu›¢ºï∆°©Ïà†ÒgZÿ,î≥ÕåÆçµï™±˘œx4WÛIï^kp4€qñqB«¿Zv{Õ‰¨ﬁ˝^-…˘¢éÒúâdÇ¸8Ú¥»BÊ˜ÁuƒÁ-⁄±èœ¸–π‰L¨¿E;-ì{7uÃ+π‰πì±èÓï≤ a•¥wº<Ï
ÀX/πW±{Ö≤VÏŒZQ!Vm[öµó¯–Òƒ≥pÜLaÎ[pŸ‘w;_⁄Âk/X˙_U¢ôD™◊M≤lW∫—2ö;mñ†#¢#ÿ2)Çªt®·ÉUH+¡á¶é⁄Õ‡#Éê6á⁄£¿Œ'o„(r‚”~∞∞åöEΩ·±´èÃ”ª¬;∫∏<¯	πpÀîHùƒ¡X5‚πäp5USEj´<˛èï òK‚‚"ª≈ÛÓÎ¸•©rxâ4É˚¢‹I’ ¿≈\Í¶©L≤®ú*úÜö–CÒ<üal9ó8õH<óY®ìhìÚ≠b	ú¨∏\Râ‚R•≈–S∞:ä5cäVMÇ™»ê≠ xi9kEZ◊:¢§⁄ãm'Pﬂâ¢Õ˛P∑\1ÀŸ™ZÆÚ¨ﬂˆﬂIÛ÷%4SØá¥ ˝«˝µ~eR%>êkÒ)ÁV˚©úöVÏVq„–9Kkò˜YÍqÚ‰¨8•¡ˆòÎ@sÿ^ÒÙ'÷{† jKò0WæßŒ≤]x5œ%œRVÑ·ËEpÜNÃ—Âºìgø}–{˜5•œπÖJôò"8œú=Á‰ÿ’¡™4ÿÏ∞/ç∑±1ÂLµ∂¶`î§dœ ÿØú1ÉÂ4C{?åßÏ)ßJ∞ù5—¨B_J>ı)¶eæM ¯]ù„
»Y´uVjM≤%≈©Ÿé[JNFF> pyRá:æ£’üîÁQ Z›7§ùùÏÌΩ:‰uó`«LÅ1 å»Ñt\¶HD˘
÷∑ñCUU|^—óaôù∏Y≠tÇ‹õ∆Rßºà~‡ı—îáU∆oàﬂZ¢∆ôtLw÷î‚O‹jÈ™òiSTËW÷ƒ˛k3€úeüÊüy>¸ë,ïslıºº¡Í£riÉº*	À†LùØO—®∆ö&Ÿ4≈äÂI4kı7+§ê7¨»6ˇST√óEmˆîß≠æ†|\%ß9ã´÷ªgI—(t4Eg¶ÄóQ6œPú;Ö¡íd€nÿËD]—µ◊x3óËòîA”$´áŒíªæ3∫5[yRÎ˚÷,ô®$4¿ÖêÔÄ˙·⁄ô¢˝MÎS≠’Yhk‘VÑ¢’’’ﬁ 6†ûH1´7.L
s†~jâ¥ô»Ÿî;˙’óV·€ºN»Ó≈ò“¬öÌÔ<+Ä9F;¬œ∏T˜â	4–^∑-™∑îôE‚îÆŒ]"N•äÇP¸ú‚QΩdùÖ ﬁL‡-›j.Ô·⁄8É¨J^o~´•xfÙÔ¨•¢bl§ï“Øz˜X¯/≥’a¡¥K4ÿµË–ã(¥0åœÔ:XM≠‰z\ï]ã?+KI√Qñ'∫¡,™\¢29æ∫=VTW{§¬´V;{ÿ∫:ñ~$<•ô©ù≤N^È†í”"åÃœwºx¢Vﬂ(-Ñ∑nX¨¨vÍÊbY±tc’MCk—<h∂&ë*K≈,ùÎk◊VVwﬂâDÊ©ö ‡ügñ5zY‘Ü‘÷˙ÿ÷§>7;X¶<◊À©1BúÓ8„Æ5ˇ/ÒÃ¸M€ö0”ó6ﬁüLF‹ø¡&Ó*üDÈ_cn_ŸFìeûî—‡{MàPé¡§çŒ£Î§Ís…b9^/mµÆIáL+’5ç±v“zsîì#ïE
¥¶M+Ã+yP—ØPUóÓ	Cÿp…X4ñì'Ï6¢jΩ´áÇù„©Xıî¶Œ∏og“Ø 4È˝1πløríS Âã6(I·' ¢eìcMÅ∏,˜…)˛ëŸcL¢∆*¨.À6}~Û™ï¬£“J≤ı¨[â[…5XºÃÅ,ÖëX9Ëø£˚ö‚V:ú:ˇ¥∫ËAFä;◊Ñ∆¿AÁ≠èÑ¶g	™2°Áùï>wÄcÆp2Z“¸?6ƒ"≠Iab˜Ω Ä6-"ìñÙÆ√z…'B–R˛êÚWèìøù	'µGëÆ™¬çœ#='(gT˜F!iØëC:r'6D¨£*64•‡¡∫Ì¯ÉâbTh~˝\dΩØj¡GUBäW^pò¿ /Í<“Æœ<»≥zpééË%: ~úˇ;Ò7ı>Ú‹BùwŒ∞’µ®¬∆vÍ‘÷|TÁî~Ù/∞«Òe˜¬KF¸6cﬂ1 lN¬—drÉõ§K«ˇÄz~[˙È7dmQ<6õ:ÎºeHwxµÿ÷Â~¡°(6ﬁ`«‘üuÈÀÌÜ!iö„@%°ÛÙ˙%≥9óDvX[Ëøñ‹°ﬁE˘˘∆>π{Î$2$ÉÎ_{CVst˝K‚¡nVê^ÕÔkDgmu∞Ú=µ moûOüæ™˙*Lº3XG®óqÌS|<™TåÆÂ◊ÁÀòH¿3#‘â√†›B:1ÖÊ‰?íLçu≠¬®}∏ª}¥wp¯∂g≤ríÚî˛+ï"LˆöÕÜk·“ÕbÛ/˙≥±˜ã#LkÔ7q#€|†≥ 	üp‹–8ıp√N@4∆SüXƒ,Ó<¨§eÏæP`9£¡¿sBÀçNî–•ümf™m`óÖ‘•›ﬁ=xK‚Œ‹O|†_(ºçi‰:Õ†m_Y⁄:s"Ø˛Õªf|˛w»‰5ˇ‚EsÕLGŒ∆tNr˝K‰1ˆ3%Û∏∫Ukki
Hqj@`±ò€8/Ôƒ‰]¥)¿>˙ÛëÑöîNF~ÿ0Am©%PÂwå—K¶›+‘,ìv∞Z≤ˆrY8*¶…’`S‹02b∆V	nlÛÉï2ª{§C∂—
roLW&
œ≈dœÖŸá˛ıﬂá 5æÓÜÍÜ_BÅìñÙÖ]·e	–l’h 5˘ ™rˇN©K[K"^∆Ú‹FP˛|Ä~@t¡!m,‘f÷(ä¡|VàÁÂ‰ú‘Ï=â	‘¶‰ªÉ > πë¶Ä*32üFk√VÃUk?Ã–wa˚xÃ‘Ì2VÃ?uY„‹OŸÿ-“M~ZËN˘TÀl2!ﬂr›à∆q√ú%“,i˘µ©Íhg ﬁÌ◊˚øˇ·∂«äR·K_Ø3cÍìªΩ∏#7]l·”¡Ìùc<ˆù14x‰çé#`±I{oÃºf|‡È~øø˚|iˇıs+&ƒ*eä˘8RMO∏‚óbqÍ˛⁄è©_⁄ô«ÇÎƒßyÙM.àK`/dÚí\èn‚ñ—C“‡ﬂå}z“4Lùl9: ∞ŒŸìT©BFReÑ°√˙Ó@\ıiÀèµ§_åú!E—x£f–πoåÑdúπ\¶iëÔÿ6KñMe¨FÈ¸N<ü.‘?Ï¯x≤‡·Dó~cÒÜ¥3ßûÎ“¿‚ùúÊµ©u·>n∞∆y»’1{¸¥˚∂˜Œ¶L«Ò€∫zºœ@Ò	¡$!œ‡ıv√∂F Ωà24p°°fÂ
·LxYÑ…∂hÿöâü'&qÇGòÂàÆˇŸäwúƒysí/üÕ˚ı»P[ÜØÜ<X±8t∂IˇIRc)∂5 f¶ib,v0 wñ9∂Òœ=ÈÁœ°áñDÒA £ZÓ¶úÓà0
1öC¨b∆£sKÑUì’ΩÜ¬6E©Ú>†êº‹Yx√¸ÊTÒº∫
ÂY‡ç Æıß¬hX∑q4xRû·ï≈È‡√—∞—3èû7;≤È≥¬„?“Aê$éWã8µT†f+ÄpÁ^§≤®∞∫˛?~‚‹ô_ˇ%ºÛö∂›ãu“_]Í≠?Ÿ≈È÷R©ç≠7¸ªSõªW/∑…m|
mõ·GÊT5õ9»÷4#e”óÇëã4ç=gè…Ñc*∆⁄¡ZË∆dàK¬=Äå√®jÑíÌíˇ ≤ˆKØ)¥M9‚4	öuê˘‰:NM÷7ÆºD◊ø >Ú∞E¿¯A8
IA«€5’’ÏπÓµ/"ÚÀwõËƒÁ¢ü)˙˜YP]n@µ
+ó_s¢—ı/5qó0Ú5AàÂÿÙ?ZDú	7=ûê≠G£®3˛Jûê›L?‹^Ùó:»Kk”üŒDµ,ü6§;ø§‡Ó<†˚>Åø@ıíè‚ùv≥◊/)Â}5ÊÃ;¸taÛ˜,ZÉT>s–I∏	°M)4Ysƒ*vÎÄ z %≥+¶A+3ˆY	/<7ﬂê;Ô„3tBÇ0ç9FCL#!$£ÃQò	'ƒaÁ>k^Ûâ4⁄∞≠XùÁ4≠Œ+à2kªñÓR…ø©	·-¯EYR^ùkº2;D~Â©!
o5Ù,ô3XHÓWo·4>=v‚è∞Hà€«˜	∏0jâ
a%⁄»ëc\Úﬁ”çàÿ	=ÕgK∆ÉåqP…´Vã@[‚k≠æU{ùfæ¿ÖU‡|aì-?e±,Ã-„hHÄ|‡t»ÚHÑ5ú^µE4ÜPA/lJ±‘O5¡‘‰Á4·bˆ`√¬mÍÜceÉúÉ’√N÷D÷¿mì_‘z˛CÄwL≠≈1∆BÔõΩ¯ñª|ùü9.UˇF,◊˜ÙçøﬂŒZ™}´VZ8ëS¯OV•™§⁄8JEN!ñ£mπŒﬁZßq“≈1ÛÀNW7ˇ v’a„ü(pÔÒ„«u€"Œ¡„Óåuœöó°îˆÇ∂r‹µ—’ı`|ãÿ#˘¸hîc¬yg˘A9Lúi∆a·FÅ€í∏™¨⁄ûºÍÙñø¬™S˝ªø˝Flµó.@ù_ñƒˆ÷/ﬂ¡’ä¥ìÌ˜¿ÿ≤'<ñ{í3πµØ2œ%Ùç’®˘:·7ΩjM(∞/øa≥‚7q∫—ÄÙïZ∏ô≈qææ≤`ù”}Cv}˛FìÂ.p“0⁄Sê∑mÚ_©ûÍ_≤|^<§N48ïﬂué„–ü 9ÒÈIÇ˘J¬qgy©O¯ƒ≤[\≤≈5Xhd*µt¡HÔgx¬’>)YáÍõ-òàˆi¸ßâá!Ê"#÷,E—«>Íœ∆ëA;btÇ Âv<Œ√I‚{≠ú76+·UTQ„Ó—Hõ=øfÚé…ï⁄iGe˛œ÷?FŒ5ÅKøXŒò∆"å1mÈU{ﬁ‚dˇâ^ÓÑÁA£ŸÚ<ÕÈ%◊ﬁÌ‚°€B}Ìw•’∆{‚ ™`ılΩhwQÃªCOúâüÿ:¸,-ûIí$ßV—»Í5·‡Ñ©(—¡i0âê	=…®}8‡‚§›¬ﬂm÷/ÊÓœ/≤Vª®"Ö6¯ÿÏ&Û)úxtJ˚¸*„P∂óOHüÌzf≠ú|Íï¯˘ePÁØŸÄ˘Lz˝¸“û=y√S¸Oû€3Åd¨™IÃ3≥∑∫˙7wıô∂S5%„Ö¢2„˚f:1;¡ﬂæ”%këØ¨%ë-±âSﬂ ÈGoº˜‰<à˜4[Wl›ÊòB5 î√zAÙêπÊµ”ÛºıÂGJ`n#¨#/h≥Rq%Ì∂sü3"Î∞…vÅÒr|∫éP‡mãbgVÌøO€ÄÀ=´ÁmìZÊW]àxÒ*$¡¥|ßiÒBºãÿÈÖ¥XúDqË1æûiÁŸ9ñ÷>¥=ü¯≈2JÖûª1“Ö˛¸Ò˙¥R¡ﬂ¿!XÉç¯âô¢€¬r“QF∂Áø8\€øaª≤,ü´Óì)ã≥ûÆÃ}íñÓ⁄\@°“Eô?≤ba§˜g„Å¯eA„˘e≤2ﬁ…~}6ˇî‡Áî?}:H÷ö≠∫ÇUv∏a„[èW}hT·i≥iß\÷{S¢‰µÊ}?z”é»±a¡»S=.:ùéÖEIÄçœˆ∞QhÁ{G◊*ﬁA¥®Æ.ã„yÍE9[w1≥& Îrw¡µˇ§5a§ö·ËM≥L√•ÈX≈Yî˚‚VÌ˜XP⁄ÆRÆ‘ai˝¶¯>*™3ÓaûA A‹.ı=%≠Ç«Ûç(
˝ò=£…⁄©fƒD•BÕvÊìQ∫≥Œ¢òÔbπÖ≈næd˛,xñ÷8È¸p–≤*ÿõ-V√›±nºŸÕe˝◊≥8u#¨ÛVê:ªÅﬂzù05m*(ÕféüªÜØ/k¯êÏ⁄E∑÷y;ÚÎ¶Œ7Eï“¨IÌ˘’0µ=øf4-"◊IäJ¢zvŸÿÔÄGµÊm¯^[ñ(O‘ÄÕµ‚ŒlÁQOÇÍπÿYÿjπ+∂sf~u6FπÜø≠USõ\hqß≠‹f˘≈´(‘Œ≥§Ö‡•ΩK©vS
YTR0çSH39Æ~M≈‚9‘|H/#˜z∏Ù∫uE⁄
lQ¡Y<%:,!‹:083≤êÈ±à†%5%“´vÃ›ô]ê”gå≤¡ΩYüf¨&gô’ÌJ«•Ø⁄⁄v+XmΩaÅªÜ5’’§^‘_+'«nOôÜcù∞*l˜Y_ÎY)•+}17elÂÃE‹ñâ≤ÏÏóUoØÆòªÎ≈Œ±œ*n÷≤∫íV’∂ÇøıÃ´9˝PYNEº‚ 5È ◊Ö"≈Œ{ˆı√$yıÅ\™ Ω˘P»_i\S·°‘À@ôÃPπap‚E#'™Ø¶^Rqﬁ|hﬂõ~∏ZTå‡WóﬂÍò–[)¥˜≠∂<üêM5Ω’bºôu°¯b‰µ®ø{1'^D^¬’J†S.¶g¨#ß¨{∆B‚D(`«¿¡h∏Jêãâ£8"dË1íqÊ#+A~
◊ˇN∆°+Èì1~Í¯FhúÇ,p˝˜à%ÆßYa‹-Õ§§û+˛jÌK3Øô%¬Õk§[ù3á)`˜∂"KR¨•æB¶≠DøtÇÎø:,˜ÉÌaıµ4…
ã£UGó("oo‘«ïËº]ñ 4≤™M1M£§Û®
∫”X]X"©E∆áπ|6˜”0)’#N√†hO"ÀÂ‡8˘jHáfª|ÏHe;^∏ﬂ›'Eÿi∏dÚë_«îTXÇÊ+ó±åÊ∆’‰ÌÍEúCô‹ö8ûˇ≠BÆ˙ÿ~˚xÌÏ¸ÒáÎÇb¬'~ä„—[<∆W±òá∫pÓ£µfÖsUÑ≈dïS¨jÎöäè@œè—∫Ç•ê™k=÷ÑÎW†(ÁcaÆÔc¨Ä”RÎ]ü≤Ë´ ∆√=Ä"?8ﬁE®yaùË;„⁄Öñπ£˝Ë˙Ôﬁ(åâ)U”[kﬁ‚¸«k:†qå…∑ÚüˇÚøya««#lUë î¡Uˆà›U¶6êñ˝ôÂ˙=⁄*∏/QÕ9§#Ö«æc• ºª"∏?éÄãrØˇ=¿†põsíjˆtÖiR©Hò=äLÔ¿ë{π∂˚ÀúUN⁄ﬂyV#êc.|nddπtÜë„z0îNv"rÖ£éÉä,F«œ<G"ÍaNÆ9óûŒo¶è=*û™J†∞áä˙Ä£ä–™„…UB:œ≥Ùb4éºp‚Ëbåº,cù8HoG◊√™L1∞åö©î;˜$&äÁ£v–±:Fkt<9)πœ◊nùº}ßdÑJ)±˙Â
ÑÀR:,F“K€SƒCÀE¸CAr∫ô0^œh¶“G”*ƒ8/#Ãt∞∞xÆ.îXXÇ‚Áò3ë EM¬ƒÒ≥zl-†0lÈ´z3˛¶ì6Úî◊üK=˘ÿMT˝ Ìbä˘°≈ÆæÔh‰?⁄ˆ%Wî´ˆdÆ7'"¢ò5Ädæ$∫Áò	-ÃûzÚﬁú)¥K]{Ó?Vm≥y˜ãòπ0üƒÔ“V—¸êﬁ/6ßk*˜2®¥Vr@ê¨¥‹4·‹˝aà¿¯ƒie^2ïÔ˜zä“{¯2U/ôpån⁄Ö”^˚¸6&§+t∑úb˛ıqOÚ√◊¥√eÇ£Ç≤¯ÁTÀñ)MÔf√>†p<iÂ(~ò\≤¨î˘àZ√Ï|i)V8≠◊òmæ⁄ZÿÑ4·íúYIΩ§∞?Ìæﬁ~±≥ß{N^¯¥MI∑Ù([v‰ˇïç‰´NRï¶˛1iÅ/rTKLl– V∏"ÖÁ|ësÿ∑XÁ¯„•¥Ã¯Õfï˜Æˇı]ÇãΩ{`≥÷¢aq
¡ó«wc•q(Â|·∆#ı:W€êWbCÎsœÏöò4#£W4i	Àµ¿ô Mô†∂¶JPvo w]]°Hc§≥cãY≠Gñs≠úw∏tl†)µ;˝péÒNß¯öbµó·ﬂJ˘eªà‘Û
{|
ú‡«N/_ä‚AÙ4GπÂ
1(ìQ—µpˆ<Õ iπ·%¥F-≈Ø8Fó∏©|‰	lZi∑Hûíçm«ßÅÎ§lqø«ÿ‚uR‚oŒŒWNj&œÛ»Wamπªféª¨z“*‡AhéYZΩÇ" xÊI0Ä/õSˇZùG‰ctˇÆ_ÚÂåyÓ˛uoöí…Zﬂ«iJ~ÕÆsfß√:◊:}n¬òé<nåIÿÑ§u{hÃwX”2Í’‰õUzÏfÚm®∏…Z'Ñé÷IÓûQ‚M%GçÔÃû§=ï=BÎΩ91%au¢*1À\≈q$O’;ÙÛN‘b&Yïv&¡XÂnH[‘ná˚í‘c;S≤[æ QÈøú9√Á
+Æ§ûD•u÷f=zFP7å™/≈Äˇæ'øÄ£÷œi:˚®ß*—ëö∆©ëù˘û≈≠sal…,Ó<µô¨cçn¨Q—ª˜qëí€”^à\cF·[¨‡€ﬁªä¸´Ô•Œ˜éûl9ÁÀÁ∞.èN?ãÏâ¨ Ê™∂ fz6V∫¢&fQß$tÅ2ì5ÚÅ#º7ï8m]QL¢N%Ü^WäQÔ◊Æ≈ öïUl3öúƒêqií mçŸ∂Œ”O‰·◊Â'‚3»{ª“gúR—∏T Oäë?¬ÙÙ∏wv˙Óôör˚RC$òò%t8â–Î„ˇπuCrË≈–ø”ƒñ¢4`	rﬂCﬁ¯9ç–¯ìª!¯·0$îú–(b7ÅN9.π,%3:dƒ÷FífV˛6¥â6õ™o¨“ñSÂõè™zn≈Ju‹F˚t·ﬁtÈ7‰µsÊYÇtr‰«‰7KWıP)j ±†2yÍKèz˘W†îŸ]U±ebª`ÿÌç*⁄IÂ‚	⁄©.äÓÄïD∞:ãQâ®Ò•πß‚
‚|¿©Qv®±w>-u⁄ àÏ¸“J%hz†1@ÓÍª’OvÃc˜Ê\7úô~âh,”èú√jŸz‚≈[àÙÍH(Ω≥°œsåí0Ùcu‘‹]Ç$,Òëk¯ìô¡ivÄRÅî∆b˜&¶Q\ó>â’3ÈP`°Ωt¸V5C◊≠ê#/8˘Ú»ÙóDé`º6îË0<éh˝)ﬁµ7¸>(˚a¨™q™&˛Ú¶ÎK©¥*?-=≥-Z´…≈Ã3Ï°L´ÀæÈúãêá\’sxåÖ‘‘˛à§ÊQ3ƒBHm˝°ÛÌVú>Â·∞.∞∫á>øQëﬂƒ$+>Âå'ûXâ7’Qj¿ÙÆj÷E…¯˙‘qôK.ıùäx#˘igºÌ∏ºîÑ1cäŸÿHx‚¡Ë=∆èo¿œa0‹‹ﬁ›¯ÊüÔ~É¯˙Ô—¿ÉôÓæ:T?ıA8óﬂ‰ .=ë‚¸{óÏ˙ `πîøÑL*å<]π @ÚÚlyQÂGíqºÉ	[ÙòÏÔ<ªœãp1ñò∞‡ﬁqŸ<_{•éÃPY8€óX‡kGtJ~«´ä2Y6W4®F%IŸ^ı?eF©HÓj§√°ZUY‚+ß…‰2>6°ãeÊ}πûBNâUmÉŸ¿Q∆9 Lí&á	IR†v$ƒ˜Œx‚crä:ıÍ~t˝oùüºxÇIõDå¿Ó∂sLØˇÍ¯ß¯-Œ AÉﬁuñÕz>Œñ337Hg€*“R'©"Ãì≠US'¶4ä˙NÂ,˛ÚûY; F£
¿zA@#Ì>":ÖcrÄZø≠ó/…
#ñ1Íé¬!h8â±Ãu‚C$jSM•Í‚„ñªlHdÀÖN‰ù∞ ƒ/1-O'◊ƒw¢E√»tâ≤˙ÄyJï•Ã!]ÎÙ7Y¿úëQHÆÂ”öú:„±4#6ÌZS dı´/_Ò
ÇïÂ[(HôK“â©Jö≥
¢zo>ﬁÚ ≈it„kVL≤L…x©K¢%òPNræ∞πıÚ’ﬁ∂^<€⁄ŸõCıD#∂Ù∂ „@⁄€Áá&‹Aßä"©ø¡-£«8—Ã˜‡B¬ZÈMÊ`•Ó4º•VW;X√}Ô¿ú^Ïëù]≤{∏øª˝bÎÂãù≠ù›C≤∑≥˜˙hÔÂıø>±Ωux€¿∏ícâY%Ì6ÚõÊ€ÂÓÚ⁄;ÑN∆ÕÖÜ©)T¸TÖ—9C(öÔBúhC¯,ºWÄŒlÈ)∑ê¿ú◊>%∏äïn ±{ØØˇı`˚≈^*…‹R}Y3ƒ"ºÓxgL®}È¥ü#`,)2g}!¨ö8/= ˇ@vG^åLôˇ9¶Úú…ÖMò-π¨F ˆÃ¸ÈZ5◊_æë≈pô«Ü˝4¢ƒêü8)ö¥˚|?æ⁄:⁄= ™dÿ»”µFæA5¸˙Zç_ØL˘NÄò±N¶ØÔ∞j‹≠eÒ>iπÓ“´WKópëﬂ˝n}42•aõ≠öh]2çH¸§tˇjÆÎ∞îÑá¯&˛AÉfyrÿ»¸∏Îê}v™¢˜Ã¿—ÇjUÙ^´V‰Æ*∞R']!∆†?Ä6"C◊’¨2¯≤ê¡ÕËSW¨®T”8 ñP¨©ÓLt•ê≈Å˘:eÀÚ8.,ÍkΩÄ•‚gëÛı¢8^l€˜òñ gr¯Óûi˛Zµö¯›÷}®•g&jxNéaóìâá0ŸÅ,&*8®å1yÅáÖbÒNAõê+≥LŸ/ÕEäÕ'XÖç≠À	hS?¶‚w@G¿∑Ïó∏^c:p	fK· E'H~·-B‡“)ñLY˘j ΩêCcyÖyÛ…VsˆHØ’&ü¨K•˚UìâB0dÉIl]gÚ?ÆdÔÇ<˝Pˆ>©íK…‰\X™∂#ï2’öJ®u∫ß¥ZY^·ihÖ0ÿ/,¿≠U·].”‘ï<∞!o)˜î»î±k5éÊ≥àÅŸ{P€…ïÀöâí¨;•ÜR·Ç#‚GÚ  ¥1˚„æOÇYtæŸÄﬁå˝–qãX∏ÖÈ˝úà§ÓÊ”Å5cqÊ‘KçÖóŸÍr™ôæí·aKŸ<(ª
Ò0EËB([.$∞`ë·læ¯ı5≥,1ùûEÔ_∞≤≥Ÿ#Êl&j;c]^úy !˜„0bï≤[ü-\ÈÖ∆¥Û˚Øüﬂ'øﬂé5°zénﬁø\tIˇ’ã}Z,Ÿ≠	™$Õ<ƒ=N<ü.g0†c†-Kó~S -B3!eÍ·\@È¸ÁÄcÄcñ_”A(k]ÖægŒ≤Å–1|nA<»îhsëJÄâ˘»Ôà‘PT-ﬁöÃ∞]–D~ìæIVCYÒ›*íj!2‰zYj»ÔZ	Öj.;4T–ﬂû‰0õ∆ˇ6≈±Æ[ıπÖàl1~%≤D‹Ö=˘&QXKø∏o‚EıÁ˘àV·vÖã‘íåylaóó»·õC¯+¸I~ÙB1=fÇOaj67‘õîg<Rﬂèoí«-HröEèÜ‡q'≈ë›ÔÿsÖÿA≈∏ôƒQt≤Ë◊wDîùg¨‹u#)D·°vk¢HŸõÌõ0ÚM±Fînî≠_)”+»$“m+°$~©døœ€MÃé§ú∆bH¡'ñO§äÚ∆}N	%_ã_âx"ïû¨T¶¸&†‘
(M±ä¥sÿ±æ	)3
)%Jvªb ™S~ú8n‰`	çW4òH¡“sóHX‹Ÿ7q‰ƒë4‚pVYDw]Ï·x±%ÃΩéI≥∆Ã€/pŒË0è¬r§Äí€7nW∆¯&X|,	%iÇQ	!J§ÅŸ≤ëﬁ≥"ƒ√3Jw@l(À
D¯â≈Ö,";›Öœ)(§Ò+ KÒM@∞˛î"åã	0Dh˛-	œXÄCHñ◊óIª,)t»ØOT∏}˘`≠Kˆù1åGﬁË8r\)MEõe¡ë≠’πÀY<“7!aﬁBÇ∏ﬁ0úQL®óíg–]à0q'eÜm'FÛÑÉyBò∞J¸8åÒhiÃfô§≥Ñyy|z◊ˇ¡‚Ü…	õ©À“)‹ö8·kﬂ¢XëuÙï¯~ì2Ê&eî°4ó6πÚÜ¯Ví8ÚªV2Gˆ¯,RGu§üM˙(EH!%ªµ™B`ÒJÅ”ÈFè‹˙ìtéãÿäáÖ›˝LRãóA≈%¸ï0VvÊõcbr†)˙Œ⁄‹û,‚J.(¿«⁄$∏⁄ñe√cq«¨B2<ùImZìü5ŒÚ◊„x…íÜ∆ÔÒ”Ó€ﬁ;}Nzº∞n>ª(
Ál«˛£©ç+M˝"ºn$”Yˇ†º]ÕNZÕ…ΩÚíq¶πöı†qVH´tì∫K£c¯SP±ÈR‚…˘_˚}-=≥Ó∞>ŒB!I,ôáá^§	c’bê~ˇÙ%bÂÑñkïÑèı)∑√(Ú˛àr‚É'º:lL&´>è¢ŒIVñÉ∆d¿3B≤	a°ﬁ abvÔsZbI np∆ãØz¨(©;É‡¡^	@Æ‚%NT¢ì.…¢M&›:ÁR™ÌµÂ+^Q .á¢ÇI‹∆¢7)G± ƒ¨M.©1≥${ÍÍ‰¥R†Võ+µhxT©anõÙ7Øz+>•„9@-Âi	ﬂAò¿∫=8/ñ¶}≈@-íJ«P]ÓTìhwgÄkï‹Ä».IsÜ9Õ‹ú¿DmsHŸáa¢àßÑ^%y'#H_ ¨l√ËíΩ
çﬂºÿj5xÎ5=Oø◊ø;éÿk÷Ô‡º83VËˆ‡ZíDÁ÷¢Â9Bu>÷Ê@}‰ã ùäøbÅ=ÊQùBCc®6Ál:§¨∂2–àG∞\∞í¿>c±\èb±qˆ’≈‰TvŸúF	pÉ„Ì•§d©£–“-ª¶Œëúø’sïiAJ‰=ë9‘õÏ2é±ˆ4	«88Úå◊d-ﬁÚiî¡nb2€tÔ◊™˘∆Sç´QΩÛ“çÈüå–!ñ±	czªFµ≤Ü>≥ôÅ#-÷J˙cjò˜!%g·‡˙ﬂID1órà˛=¡ı_ë∂ã»~r∆zƒﬁcºh»Äêﬁ'ìƒÉW)röéTƒGh .˛ƒãHHNë›˝{‰XÍuˆ.F(å#ä‚¬∞ôj˛'1>y¬jAùi]â¥≤¸ß´˜÷¯LŸA¨¶¨‰R4zÅ‚~-ùœ_;b‚sª%¬˚0/-âºÃkw[áDfîòÛëÅ˙úàtú3ª?!v†Ô¶¯≤î °&ºü°J‹ÃzVÌ£ëÓ ’˜€ÿê’™ÛlfñÕı1»œXÎ#)Kë©≈yM	ç~Z∂ò∑ÚÅUIÜÊ⁄–¢≠1xÎâuƒ‘„?Å∏z˝ó∞a^z>ç0’z–«"ÛS:a5(ó˝]ÿ\Óˆª=ãtıw`°é∞^+b0ËààJÆ≥,ù 6YÓ¨Èúñ∏0Cá∑E—ïL˚Õ∏°ÖÕ√…8å∞2Ÿ°1ŒBˇÃc«ª)i}çÍ|ïÜh_v¿Õ°„ü:4~¯¯øGéÁwÅKh~<7${∆g◊«,™éÊY Û—◊U éQcÅÃ€ØziYﬂÚÌ˜À€Àèóæ[ÿTiØTï-Vë,©õñë|˚˝≥5¯ﬂÍ;S…OW"“æïπïM*ÿÿU)¨S¿€îS¬(*"ÇºË˙ßŸúP“@5:)Ë⁄Öã∏]TñXê\Ñ©J#˙•äeSÃY„≤¢9ho=⁄Í?ﬁ}g®ÈÉg Ú#èÇÒÏ‹ã<¥ ∂©º>◊s√E≠·k#b†¥L*DQ≈yÖ¿\@A%π`0â◊±Ú;`áÙ%C‡•Â^Ö√©8˜3†‰…4'— CLï$ì€0iäœÖw⁄ô°í5©@jÂéô˛ÜcN`\3ß+ Õ?QDì∞€Ìn,Ò«U-M'⁄“3⁄“9„våCUªA§=§óO¶¯¸;˜¶Ò’á´lQÆ6·?SßãäiÎ√ˆ›¬µ>MÔ_+É`f\* •ñ6∂¢Ñ‹<`UÄ*åTÜ.≥ntV◊RöŸSWT÷æ€XLà}!iøFH€Êpˆu 4«≈f}TxÁ∂˙S†3¯@g`VIVZ•’5Ë◊nƒBë2U·’TK‡(∞hÌ·⁄ ÍÓªäÍâõƒ}Ì>‹]€≈◊åó2∂ù`@Å•T¨∞NØc≥6‹…•∏(™’pΩÿ9ˆ©˚dÍ≈¯4
~˛ô|W<˘ÿ≠"Ó‡≠“C®"™G0Ÿ∂¥Y8èAÌw£ä∂∑1GÙ°RÚ¨¬38ZxN]Ê)[4–˚Ÿ≤-u,ZÁªaaÑó¿>-ïè1*Nˆ	$˚Q8†qåÆ @â4R∑íFh+ıÉ€>•Éè€^4»Õ%|B•ró∑Ä«l2µR„í‚tûQP— 0øÄo“ˇW)˝ó}Ê•‡æ$øvÈ_JSü.o6ÄÏ\”X Ì0W≥ó
˙*«∫.9Jµël¬ÓtÑ ˆLQyû±ö»¬~Y‡øH}Ç¥¨)`"≈ãTx˝:Ñ
Œ"‡LqTM¥È;∑%Td~úvbE:.V©°¢ ©.à¯¶$a∞Øõ¸üojÉØHmê"˜WØ0hÜ€GÖwn∑ÒqÈ%∞U#’7—mcπ–‹5,ø-mÇô˚jÿÇ4◊"§ 'IÿÚ≠“CπÑm¬:ÉAN¯)	7ùfCEÇƒ√~”%|µ∫ÖSˆ7•¬V*X∫©ñ‘E/˜ä!√{YÕòÊnËö(9ùØåŒµF°51+Må~Ô*Õ	ÎF/ª+ìä,l¶qq»"LFa1
.µù2'c™t/2Òﬂ•õÕÙ<™ ã[Q¯î„òËT"…Õı?Uêüü‰í√Ñ•‰"s”),í%ÚÇAÀ'\™YdUóNäëjTdòrΩ^§®¬ë,$µ;è ÊC]Œ⁄ê¢»£v◊F…mi‹ yßHWücöËëªYû¿
µ=¨◊R.À¢=œ[√0ÈÙK7]≠E≠ø)VK—éz™ıW∏Œ71€R
›n»{1—¿±¢Ê&oﬁ¸∞¯4[dé[átîü-ı…ûDÅ6GQVÖ÷;#‹ù§™iÊœ~'Ìi>ü+]í,”:/ßö«’\‰ﬂœD7_ágNæOø&B)aµùHKÊ≠í9¡ÿ‘ax˙Ç@Ób¨ù
—
n R¡'	ß€&˚€T˜Sô€˚˜Øwˇ˘˝˚ÖÕﬂ§qÆ-"E¿3ŒØ2iwôGºà¡ËÇ^$IH~rÅÁ>Åøù5Ã"…:πÑ√~G±HπÁ3`¥$X£4
bâ⁄ƒòYÜÁî¡ë®„°∆0Rz
ÿH£'ªÎdœÖ	Ñ˛ıﬂáﬁ ºO^:ÄÔNÇ#aì’ç§ WçsV«ö)pS"]ãñx’PØ<A⁄ZØD¿ƒOÛ£`‚â«j≤ı$	üaÉä_ïr†J5¡Øjhò*"eŒ*ÿ&ííNÀÂÁZmløô:6ªjY¬Öï™”œæªFımKÏ◊wJô†˙KÕ!®Q ¥≠lﬂ§ñÕRÜ5
”Ùá•8~®…¥∑sZÑÜ*€ú87—ÿV∆zõ™[âzË¯güMo+A“}IJ[EêÙLJ€µ_Ö“‘®¥U•ÑXNs0g>y¬Õ∑è{gßÔÊ¨ÊUkÆfKYUÚÆdÔôÇ7ØUÔ™rN(˝∑nK—+q¶·Î≈òı˘ix•îí?UàˆÔÎøÖÆî‚”™tU)ö™tÛ|Êˆ*]9c3ø˘NËyIñ‹¯≤ÉL(G(g‘qê‹˘œé]ˇ¯#˝Y¯WQOG†Œ”biö)≥€)PÚnÈ≤`ÆÅ¬¶xH‡órämz *QßúJ∞8ÍFné◊ˇ6 π0p±`sFÙapÁò7§‹ö‚N HÅ¶áqöë%/0Aña–u“0$<Ò»U—ÄEËNXá7pmƒlG€ *	∆ZNF#§jB£öÑÑ+"ˇ$[3k?1Aªoû‚BóƒYEdäJ i¢}Zu”p‚Ë∏í”öÚ@VñV
’∫h‘k3≠Q“E°ükKpó/ôÆ‰ª§ÀOQw+YL”®®∫·¥∫S[¿]¡kSGÃ≤v©#æîÖ:†öx±«#°ÁºP¥§g∂µ“”ï=¶`≤¬‘â®°¨',RŒπ-ΩØ8aÒíÌúØØˇvùfú¶âQ¨êP˚© JySóc™NsÄúÄ≤Bâ*’`êŒs8 ¢DêÙ{jãï[dNø≤ÅÍr{Â5D˛∫ûîÀd6ÌãÂÇ	˝r¡Ñ¨éBKe7≤‘‘´g·ÃBMfÓÄA%_ù˜\!™yrÄí;™î,ñ≤N/{#ÿ)p’£§≥\¢$≠j∫a*ç©íe”'∏/Â•2gµgdjkL'Œ≥{Ân0^¨/¥sÉ;÷I¸∏‰VI∞Ñò˚ƒ)ﬂãÔ[óÊ±i-N@…ß1¨Epx©O˝ôbKLb«wã±/ÃNèC@r*£—Ó®ìK≥˚‹
·r˜=å|Òã†^˘høë∞yê∞|=Á@«j@ÈK!fC#mr˝+1òeSs√ú%æÑ-	q§u&R Ö€rÚç˜†9!˛]†Z@^ﬁ#	~èŸ}Ô4≈*éÙÎ•Vˇ  ˇˇÏ}€nIzÊΩü"T=0ã≤xñ‘ÑöERcí®!’√!%Y)2g™*ÈÃ¢H6ó¿æ•¡ñoX`Æ|aco˘&˚ªè‡¯„ëÒGDV’Íô…FwÛÁ¯„?~ˇ◊†VˆXNA©ñœ/ÖJ‰¥áDA=æÁ˙©Æ$^‰0ªˇ#–@Y6‰”˘üólΩV‰®¡ìAÄ0√FÊ] uIÈA1÷¨ô§YLkJº˚ˇ≈bïÅ+cz7÷√Iò˜¶WW«√1ò8Õ0ã°z6¬≥¸MtÜÒÙ[∂Xdßœ©L}*O:¨∏¯‹xÄ“à∞«åÂ/≥[úÅÒÏ˝ﬁ·˚=¿’ﬁŸ˚Ú’¡õÌ>◊ÁŸI1†+°BgÑ€å¶&rtöOÇ:ù‡:qú—ƒÄOröQ«‚2£–ﬂça52^Æ˘®î«BÇ≠ÌmÆ$z¡(uÇ≤h*Øæp¯sﬂtOo¨
{’hìïmÅj∫th‹W«ﬁµ⁄D‚˛]kíîªﬁÿ#Û∫ŸY‰˚‡V“ıo≈Ü›˘ﬁ∏¸à÷%Zt
¡œ∂>”Ø∑˛Tã…>ﬂ˚˚ù◊?æ:àsìçu¬r˜,ëª˛Ìk˙'498ˆR¬˘Át˛4ú≤o<º3!πÛLxxü¶∂b2èo–±âpe9§L€g0[W7Ñ;6ÈcÓŸ6üÑw†ÑüÊ∫LÊ_´è=NQ…nQ≥åe8FM‡’⁄9äê¸∫∑©¿ÔNuÙ˝∆Á´c28€‹˝≈Ω´Äƒ6‹´÷I=‹tº¨L*€∑ >ì|°⁄GŒ∂ÛD
u!È{–|ÉÔ*Æ—p+ëpÑ
´œõ˚&ª`öÔ√lêëÓª≤g>b‹¬JÚ¨Gﬂ=y˙dÂÒŒ1ar∂)UR^ü2$v(H»A◊•ÇBöTür ÃÂƒ–ƒßªJM‡¬›‹ÒiºÊÙfà¶&q%¶ëÖ/MøOâ‡Å<~P´ÑG-,÷ßU9úd>?EjCÜgº~‚œ◊TûËMÑKÃàÁTbî+&nü≥ˇZ3j Äv∂Ÿ0Vzƒ\Ò∞Â⁄lµ∞'–jêT‹ÏÔø9d1SL4‹,¸bÒTP
û‚99ãOœ$‚\aÈ∞Ê≥´(xù·E>A9•jìU{˘ ?ÀˆÏe©◊0Ê]vFwùiCŸÏ>0ï 
u÷»<`g
PyüPM3ø¥æYáéX˙fô)œÑeçGﬂ≠ØØ/ØÌcπü◊Ê—⁄p™ë“¿ìeK{¶∂|gÎuy
Ë°T8t;Ñño8''	ÿŒH¬älx˛*©Ø’ëë}å)ïTˆ†ãT…%ÈÓıãÒ˝óœ˘†õ2E‰ü¥'iN‘Jëœ[{äVM¨˜TÎ",lõ˛Û˝±¶Qç›kR≤&„™Góå8d≠ºäQŒ†ÍÉ"ŸÑQmÁ≤  qaΩ@ˆÜyu∆Ω-ı.ÓÚÔﬂ7*x∂lF‘À⁄∞˝smƒÿäÑ¨ı»Àb0∂`p∞L∆˛≠áål§'1ƒ√∂'©^"›#∫Ët–¯1≤·‡
ÊSyµY¯AÉZ¡.ÒM„∏8;áÏ÷«ﬁõ¨òﬁ˜`˝üËê¬‹Qèæ€}ºª∂ªr‹Ÿ˙ëtq≥õÅÈΩ ˚>Êªy=úvÿ€Êu#æ·çÚ˘ıp” ¿ví/A@¯n≠≥ˇ≤†§ôE±ƒ¯]Ü∂≠¡äöLß>™Ë0îUÌ8y¥üˇã¿/Œ(0gÜìÚÒuùóµwOÙä—È‡í“x5!Å-“t~	(πZH^¯∆F*!ß–“S¶Ê◊µÖ‰:*!ê|@˜Ú‰-bÔo’Çjƒq¨ÅßI“a⁄÷È±´ƒIΩ†ıué√,ƒ8Æ@Ím4ƒ¡!ßJX£€Ër”Ÿ“;:r.‡3îçx®&}4Ñ_H›+……ûzO™€qÅÜØ∆˙ô±tOõ˘”’ÑsØ4ç:ÓH‰w%≤u™¶ÔU‘F>·]©P≈›Ö∞h=]>Áµ…∂)xÏ¿{Ω€}IX´∆pÉ)zã·Eï◊ﬁ∏\„,_¶S„3∞?çŒ∂,™
Îïﬂ∆	2$9(…¯˛è„Àª≠À§|áöe°W¬Î$·¥∏1çÜ‰˛ﬂjÍ®}DŸÜ∂ª A¬Ræ≤O*-3ä^êA	…Ë©HYÙ3Pë”E8æîé?•N‹0$˝/˚aÛ’&Cb?ñv‹‰îÏ‡ÕõÒµ‘¸⁄U`ùKÊ∏ÓòBp±<î›ò Ç´¬7µ‚K∏”…ﬁ5db∂úR ãáX ˆ\ ˛=Ñ‡±9ıB⁄ÖΩ_Ç{ò©kóSV —Í±πV¬È¬&_JÎÎËqÑ¯Z˝R \Ÿ”M√…B›kz[§˘≤¨%ö¥§Ò“Ωˆº˜Ü¡kö)€ZøÏ[~∞(ÁUE û-5\ƒÓ¿„7wÛqVÿbÎ¡mp@ãàœø@a£®*k>ª?nwlï_(€C$€Á‰¬êX!Ω¢>‡∞°œ…<ús@Ç›¸s9‡n–§…|v®}™È”;øN.…fØ Œ÷w—ñˆ{5≠7Ô.n8^^wdë‹éÚ+≤KWH7RLüæﬂ3N_éA(ÍŒQ¶‡x·‡
ŒÕv8(Õ6uoŸ˛›$ÏË"ü˛Bìöy≠∆èa©?}®4giŒ¸ö?3%Fÿæ≥u∏˜~ˇÄÓø~µÛÍ˝ˆ€˜{~ëf’û }Y/”è,‹ö2	eÂ[ùA! »É±Ùƒ¡9âU)˛üW{C˙,}c¯gTî	U,Ì4”ôö®âNÊ8_÷dõä1ˇ¿: é€èH8e∫VÅ˝Z±7$ﬁ Ø¡>∂xsY<”ªΩ∑ª{o¡£LıH˘õÛ?ÈÇ¶‘{∫äˆﬁ|8‹{∑}∞Ω≥Ω/+ª∏¨.Üwª¯{’mø;ÿˇ›ˆÆ™JÈS,Â∆ÙΩz˚˛`Ô7?™·3&∫4aÍ ˆv~<4˙‰á3õ∫ö›ΩﬂÌø˛ëN”á§≈L<5√¶à˘‹éÙìôV+Ü|;0‰hÖ˛Yï›Ë˜Ÿ_è}ﬁ›wΩÿƒQb⁄`{oøqˇœ˜ˇ¥ÏóOù~a€¨¢¯FdSoV¶ÃÆ2k¬Õ ÙÉ…+ÛB“{Ù◊∏[πu/ô)èJx=˙êgu9ÚÅpªg1$É%m¡‰¶dØ‘Y=èﬂî‡œ·»©è≠ ∞¡ÔRx=êâúéã˚´¢<©ÛÍ3èÌ¸≈O’Dœî5˜˘XÂqSﬁ¥üóJÊ§"≥'xª≈qpˇ§ï¶”bÚ@iˇÎ#Îã—~´Er˚d*∫J¥ÔÕ—Í:L¨R{ÕâXœÑG´äLıù¥ﬂ/˚A65áßø˜0|ﬂ7O^ÔπõäaI«ÃÔÀ–Êÿ6ô:$ ?Î3ŸI±’Ç¨mÅX¡^CÍEäº”R∞ôú]vYdqG	>}ªü'õ|zÁÙπtêü^÷ôóDm√ˆ-8JÂEòhMÇpN|KîtπtÂ°Ëó_æJ8Èláˆz∏…~WÂ7¬	ÕÌT∂∏Pb&F!^ñù€º∏ÄÏLŸ` ˙…r∫"ñLÈÑR±NU!‚C—ÁôµÇ√ıºWÙ±àTÊFUÑΩpx€áŸ¯ÙFAÛG[é‚ê¢Â¢Ù∫≥9æ#ˆ†YÿGEèafúé˘≠«]©…J"Ò*è!j+›ÃÀßØíGô6)–"ã}-‡œ√'Çg{”ë√DlWTfÏ}™ aT”á˘∏kÕõbÅ{>å.Å+ôó≥ˆ¢,y6ö˘3Ò™ŸÁác:¢v+§	qã,áf‰πÛ’ÔÀb‘ù[ sÛ°Ø6yJ8ô√åı§ -ΩXºOÙO:wÙ>≤XÊü[ùgπŒÁhü—∫Eû≥êóY≈¶SÂB‡ªº§&≥(=˘˚CvA?ˇú≥û,áæ6{≤©f˝‡õ◊;d∆Ω7en>ÉÄôÛï qﬂ<”§æõäœyu#j £:>˛|”nFè|ˆ>†¨Wõ.u∆ﬂWd!Zs—«K˚P¥Ä>†,wu‹‰‘r™ÛF¸E¢…æ'w0¥vQO ˘Ó™ÃJÒïç˚9¯|√¢";Â63àë…~ ¯<j°∞È5„aa\RûÎ]URäû?O?ˆåSØïÅÜU>ÔÂ˜º?+[ƒÌ>1™í=¥åCjÚL”1Gè‡ÿ∆Ë£<ıÍ?ÏPªl>Mbù≤Aïg˝>sîºˇ ≠TTøí∑ÅÍ«N‡
∆{y„ÌB‘—±Ë÷+QúXöü(ÃAÖ?†ãè‡áÊ^djV}⁄VŒqãW€‚T-(Ö5
QS˜(Æ&9ã’ËÜ^f◊øM~Y∫\±·}≈ªc2,ˆË-…òì@‰)~úª”∆ÊHˆéiú›Å„›XÈñWX»≠∂◊˚2´«›Œ{∑RY ‘g~ñëﬂﬂŸlH∆ »R~”W>óÉœ†¡tHß}*;—•ÉvÎL≠ú/XÕÆ¢ZﬂqÖmwÓ-·„¸sN	`9óˆ≠!wÁÇﬂ$˙oDJ–5€üè´À|A≤/1ﬁg¢Qû>A+|g hcå9–Èj“XQ«"+óÍ‹0Hêy?Áõ÷m_|‚{¨Qì"Ã·™j⁄OÓUı`ÃDCÈñ¢yë<E‹ $î0èZ+aºŒ9:¥˘µ´MÑ )ù8DÓåŸÍ˙ Ù|A[´¸HwˇÇ9÷ÊÌ‘íÿÿG∂„¬›E√*èÀy,çbníƒãTjÂ™gÿâΩÙ(a·"'¢H$`ŒïîQñ^(1øû∫íQﬁŒSäÍoË_˝Í FPˇ∂Û]JXø,?ÈËPô !Rí°⁄€]OwÇ~%ŒF˝îó$Îz¶0√Å‡≥D‹S?˙@*ˆÄ˘TºYhÓ—D~Wña≠¶#ù:À•€1«TøsW\D]Iº‚üèÍ≥à6màk¢Ã*ìúˆ6U˙û	òõ◊:[€|YeÍ‘êΩö¿)ÂRqXaß≥,§âr	IAy¡óÈÎáyVùû[ ã'5w9î∫å«Â≈‚ “*±x∏°ªæÆ·”G8jÿÄ¡^± 4ﬁÂ5`y”° °dÅqôb4≥∏wøWVªK£s&2≤ÑN±µ%£ÿ]Ïe0Î!÷ì(o˜˚ k©ƒ3î4[_ßaÆ¯!›ëUJ?ö≠T ›»jÄ¿QÂeÑ`Ë∞ãBˇKÁ	B61t¯˚,ëoì¸6ìHœ» iR|-< ÌW°ÁEøOyQ∑±∫¡ËM pC_Cq€çÖ¬KM“`õâø°Mº®7Åe§çá(MPDê[Æ¸ê  h}4*¸˜›˘Å‹¢∫‚ËÅ∏»— Ÿ™Ö=Ü˜2∫@#J1Ækx‰ÙÂ®` §„X?ÚtÜH+”¯Ω/z
)T+˚ó ¸:«G£ ÖK`B!Ùw!Îü1>~LÈ∑Ú˚¸TÏ⁄∫ÎÙ3–¬%'·ﬁÑ±ø3†ƒ\Dî∞¨◊Û¡}©π√ÁÁÉµ÷e5Óv≥r¬÷f∆*ËX8®∑)/⁄=·µÜÀa¡4t«oÑÜ÷Ñ¨dÂì“&˙"å$¢v
YÊd®[oÛ—9§Öê«S>ÇHm≤zúV`¯¢vC’–´π√pJ‡‚Tp˝3º=·!á8î∫æº¢]-{ KÊ})X√qmÅË ﬂ!é_&Jª<ıË∫Ö∫√ÜRÿ—§!›èT1h¯√&*ã‰1Ùë!££h∆Å, wÇlÅ+Ñï†ﬁ	x§Zôg¨4&>Xº`©Quâ]'*3‡7"Ò∫ƒ+ÔTÓ–0ﬁ-ùœtMaå) ¸¿UÎ]d˘b!∏™ÃD?A»C‡êx7[[ -pŸûΩ£Svñ[ÇD+o\ûñ,S™M¶Ï‚7•„müÕMÕe◊√ú“T*cû√À˘˙:j‰!.ö>àR)md5"ÏçœÛ¨èH[„ È@
Ò§™fB–ö™Œ:!D≤©´YñAyHñî‘˚li|>ÛZÏƒﬁmŒo«˝ûX5EüÎHŸ-µj2à∆õ‹ ≤n(Sf‘’C»º	Ä%{Óõ6;÷Zµuä]2úHæı4…36øØã¿ÿâ¥ÓÎ7Né›.-øïÅÊ—Gï_GI√≥ÒIŸø1õIi]—ß¥ûà1Jp;S'	Ê€»]EÖac…óñ»N68ΩdCÒÅDhπÄÎ… ¯ïYÖsv‹ên]3@ER	Ÿ'ò0ÚW@èÙ† ÎäH\rúg#“x¶tÀ≈Å˚‡h¸7ÊÖë=‘fˇ>Ñ‘£¥ıvwËÄ”øÚn∑æRÒî+Å®ÄÚkC÷^àV¿‹?ÖœÅ·–ﬁìÃ)@˘NBó§7B§ÃÛ¨~s9√ı‘iºÚ\¡ BnÅ	ç‡§dÇkr—∑x¿8≈ êﬁO¢UB±÷q^˛∂±nÓ8ﬁ[œp£*8ª∆xX 
À`i<+Œ‡ì±ÁÀ<£î7ÿ˜ÈÚ≤GÔòÛ%ã6A¨ÈÎˇˇÎø¸Áˇ˚Ø!∑ﬁwG^3J3R…n„=ã ™À∑¢îi)ÖG8?Ò[xàLgÍ2Ù±d,¸•_Lûyµ`π,ﬂ!Ì`∂È<êûÅ˝iFbRÓ,º!¬„+OÚIˆî
!17◊;JD=Ÿ^/ø∫5Nàg?êÂ@tí	ÄF#9c≠KÎUXÉì6äU(yÜ≠ò1„=·5#jæqjâÔïVB{È=£…çÑud–Ô[ÇHêW£˙Ú`ÚR˚—a•©3"õ!XG;·)∂«u¯–dmI7:?ü‘Í¸<dv∆‚!ÕÀsí¶|ké˘n&º©L}C≈Ω8„[.m˘ÀÍõ—)ˆøt∞º8ã6 ØÑõ)eÕl#a#X_`p¥áÆ+Hh∑Á f$ÈÔÁ‰ñÙzΩ™X r6Õf›pÄç·&Îk\›$è!ŸUVåÈÆ˜¢›Ú¥€áO»úíu†`ˆ·Õù_†Õı∑3πÖw`e¢\v7Ø™8¯¥æ`∆ Aﬁ£üïU∑≥GˇG≤íãH≈OBx«/hÎfgÅ@˘…çJz/à˛l¥‘Ø ≥`òÉ*' k6¯˝Üeöñê”oXZÃ*/Q	 }hÿé¡ØîsêÓÚÇEøÛ›€Èlm_éÀ·˝ó1‰¯Óæ‹{πO…ª™Êù82>f‚¸≥%˛aZ-.”
Ë$f	Ú5ì	@'L˙ÙÜ˝ë÷
∏ÄMﬁ$Ù´F$‹!ï∆·È‹È“ÁÚËæ#ó#:ø£ΩfﬂÂ◊Eu˘ËsˆŸ+:ã„ﬁ‹]»¸gı§’¿˘·∑õer:ùpvSÍ„óXÅÜ≠¿)÷p«∑á*ı‡pòïÜÊ}„…∆⁄˙^L~îóò∂p€Ó0…g∫Èlu≈ú6EÁ8ïEJgq∏Å»Ä:Ü(jé)Êﬁe˝ä∫$6$ÚFZCg;Ô∑ñÁ9˘»ó¡Ø¨˚ˆ¯ÕÙÚ;se{sˇÿeÕ%Áyê'∫ûPgÒ´&Dñ"¥É◊&Ï‡¥9äÙÀ`©≈ƒ_[Ù÷Oñácáˆ≈Ω¯"Ê+—∆:I,ÉÕƒ7Ùm˘¯â8x⁄–™fUùøç]GøHÄ†}Õéìo≈&kg;^æ5≥y˚Òjquÿ√5ÿA´f»`M≠òã»’tÿ¥¸?Éi,ÕQÉmôöåd«1ZuÌ§v‚?∫ëˇ‚òG&_ÓC<j‘$`Ò¯CÛ¬≥`öW–-É_	zìâœa¥ÕÙëAv˝7p:}’√…#‹µ‰]&‰[Ê~K%_⁄£ä±$dÇA" ì)Ü≥…≤œê®G$3’˙¿ñòˇ»§∏çwX!ÕCq#,ya‡û√H2§¨à…/êmé‰Mn‘ñπ˙&"WƒH°Œ/-S	rW Éú”](”íÍ@[·á˚µ#!‰∂VVS}¸D∞ﬂq<|É°—ﬁw;qºÕØ^ ´ ùÃ’ïS¸ÌQ_ –– ⁄8ºŒﬁﬂ"x*‡k ƒ‚_qêïab∫îUú◊öBw ≈‡NQùÔ¬›¨"/≤‚:K±Ô∆\<√ƒîkÁπ\Bπ{¸k…À3Ÿºo3êú ŒÙ1˘ÙbXàc.r‘5Q¨cá‹Ï`ñch⁄_˜º“@”Y¶oˇª™<ÕÎ∫ùÈ“[vÎ÷aÁ9…º«ü/Åï
Ä–à&‰pÙpDûO4áÇ.µöƒØLöbah_ Jà|däu!OHHª2’	â/$dp"ä:ó◊ïãı,¨‰!Áãî˝ﬂPj≠na#I™óÀQ⁄$cá+ytƒw´#P∫>cÈ±€+˚‡°(ˇ*3èÅÓwC8éÍ,D®¯ªíãnwÈl)3Ãˆ7 :y"ÀnÏ§Q≥‚≤”(á&5Ì¬ã=luôq◊;ÂËSQÈ6⁄8]9Á∂´ Ø=DìTãG≠·à~h ◊LÖlò¨Ä¡)O=-&ò¿VÌ©∫áÙƒ¢°¢ëJ)EÆhÍ!&>TΩÅ0´∂Ä)`ò¬»√ÂìùÕÔ≤∞ÜÍ•ÃÚ,µM≥ùÂ˜UVüª”|©•'ùg?±2ü˚¢@≠Ø4qc‚¢É—ß≠∂H£ ›æø_^^zÏM•wVâT◊Äy\≥ª)˘UMpùkıq ¡™âÎ!ì≠¶$X5–Åó3>$t}-=]6±;)Y°√¨ﬂOâƒﬂQâXW=âXâ∂}åßd&og@8y¡iÇÀœ!¥¬ü∆qƒrx˙†O™‚Ï2ê@–craXOŸp®5dâ—j=\8·Ú„Ê”e«ÿﬁÿHès¯àäÜıoÛfπ1†I/k‡`h5'ˆjÒe2Úkƒïeì.X4€ù?>£	π∫ÍÕëä$¿ÙË…ﬁÎìq%˘–Ó÷KgÎ-¨Lπ
ez^µL§Qën8O~ëŒ÷a(‘¥íí.ÿ|∏u•R#Aƒb˜P¶-e€¥ƒ#^€]no)µ;Õ!îqP^é£ÀŒXê˙õ˝'Àå‹6Ò‰ÿ@=1 ÃæbòË¯ ¶ún5N»nª‚ÜF‘ÁU1˙=;Üc·JYä ∑êæ}‡HE—Ÿe…‘v aI>]é^…8`0_îœìsq±·Öêq @ÒÈYÍµuÓK¯bâR°©£†åkìu†d\Iˆ9Ï•ã|@ˇK*=ü CÜl_‹ˇ¿dLW¡æ[ÄÏ"‹—†Oi≠mD+ ˙Ùn€`îÒb¡aèKßGJŒhÖ,Ü¬»πd÷Ûp/”,˙óe5$/ã|–Ø7%*È_ËñÄª€	N>Ê2Ïo∫Èô◊SìêáQ˛VúËBáÊ Íª j¨7' ‚M"˝q:[Éπà‡`Åúb˘$·Ç÷wPßSq<h’	Pq!Ù3¥©K◊4õ‡zdπîûyIﬁ-≈≈ﬂˆHROÇ†ön∞ÆüÅÀ+2ÿ° »7=Õ?CNÏN1‡d]Ùi‚GõO)U˙@.™˚ˇ∏¶3-ƒΩ’¢8ÖeŸßÏ˜'†7¥Uûõ-äd1+#á*v∂º∑[À¡Ÿ:[¸ˇ¯áòﬂø„€†06-]" úŒÛ&zN;PwYú3Ö∞LÄÏ«≥o	‹Ÿªﬁ‘»™uyRî/;&/¿ƒ—/‡º+d–?)b%¯ªπJaø∞#}ªﬂ'◊ÁêÌ dóÏ35`îKKd•G^Ê‡Ã«ÿOrzYUtä7îsòHn∞Æ< \ö˘ÃMVfPiK6˙Ïu4^ËàêuÖi˘-!Ä•qsn1≈À¨ﬁ9-≤9é_iïa<ùü˜÷j^:èõ‡ÜtÏVÂÿ©DN2˝Ç∑Ñèôz˜ ˇGëræ–¶fÈ∞“¸ÿ}•mØ$âè*=∏=†Œ∂”á™eF9´I‹sú7Ã'˛2Dê»õÏ¢	SjexëwÌ.ÚÆ«˝úªçèd
˝HáuÚrwÍ¸Ÿ∞-	DS,œÄVÚ±Íùgu◊≤>Õ˚„AìÛ¯t—$A^Â-oä,u«Õd‡´zréúL9 ÷¿G0∏®}uÇt>&ÛÒ,ñ@"=œ˙	º≠óîì√M’É† ˙≠çÅ¡ı,è_˚óçø¬§¢’
£E∑\zÕπwö∆ã¿a±+iç<0l‡\›I≈QUõ´è≤ÕŸ§‚z1 óìMr·›vÕû7:£Ú,±∫/z™Pä%z¨—sÄ
˙ÑEWÇüu-Â _p .Gƒ=U≠”¥®˜¯7]û”lÃz“”ùPÊÑpÓNÎAøºr·g%¡Ç±Ä>˙	˙Úõ&%èSwc:0“Ω@;Ù≤*á¨∑œ7)G«“ïjäÓ¥∑96Çà◊ØP2´wGŸl,˛Íñ}1÷1«∫ÌyﬂGF¨QoÜﬁ„#ıQÀ∞£»ëß‡7õë‹KoÁÓéΩıcÕ?VàbûóçY4’òû]™I	‹çÔ¡Ãˆm{5k}|÷ÿIÊL=ÈD∞èåÉÜN†ˇXK˝=ﬂ¯¿æÛü&ÊBpø	¡óÑ˝~"Œ?Q7¸ÔÎ9Ôã´pÈ\û£ÕGÛΩ§Dµb_ó¬[ì“,EãCÒq—†A(	ß´[úù¿B^d£ë«‹$>líd'.‚’¥ÿëÄçÄ@◊Z ;‰s4Ò◊Öªƒs•◊VÛˇSÖø§ı%Tæ∆%aTæÌA^çﬂWE6:”nÅÙgÈxAVˆ+"îq&`Â3	ÒÖÙÎãõu*Œ”Ù§–h( ±g€jΩK$8Ñ¯BΩ-béÜÒúÃ1êŒœ|¥‰7åà;MîfsåBaì¡Ùp$´≤;øv©ò{:ŒÜt∂ájúñCPàqe*á¨œ¬K$Ri¨;ó†Û*+ +Ò¸SÙ˘I9f∂ôí≤>≥äÇ·ÜÉôöômÓø,ö&ˆÜa¿bÄ?Üπh.B£x~ı#ø|3|4Ωçµ˙‚4º®∑yöŸW"°&%‚}'«fØ.á9á∫9±N+∂¬Du<É,•ô'='W™ıV üÆ(Ä®,¢£KΩÛ˙óÄ≠ï9á¨ª∞·k¢»¥ÅI?dø´Ú
~sp≥ÆÚSâCÆ•‰àXWØk«¸ﬂŸ∫5'#%GÄwõ
Öπ£%7Uùò›πyq67Ê†,ÁF¸,±†H‰ˇ˛œÁ S3«kPßÀ÷-Ë¬vi}rßòP9ÛΩq˘öeFÅ7dÌÔãÉπy´Úﬂé˚d`æR{gœêÇfJ≈∂I…¯ê{üíõÆÙî#pYI?läHL n^
)èì∫D|ìa◊¥#î˜π˘-g›ÕçñˆΩŒNÌ.Å¥Ôe∫Í…æn‰Ø÷›O¯>òVIW—ÃŒ|‘Îı‹„dAN€qÊà7y≥`˝Ãh	¶…Î£„•“	Œp˛5π ])‰ái⁄¸«Ñ~$ Øò®+p≈c(ãßfÆ&ÚÎË?ÂúîÑÃ“‰?íñ3«_—>èô]ßû^VuY-é 1¥Ωº¢í!bFLƒaŸ‘x/V<ö‚dçÿïO0éºG]I@Ïm!ùøªˇB∂ï„c˙UÍÆÿ=˛§√Êœ¥√ﬂä;Å’ÌG„àîä>¬c∂«Ÿ®üU}mI›ZRD9‡ÚM´4ïËa≠oπHTcª1G§~'œO´Õjª"xÒñÕJÕ¥—ŒïÛ!HÁ‹
Ì$WÌk!ú˛dû†»‰5S ˘ág«ôæ"‘õ‰Ø¢êá"›Ô¯Èéoü#≈a„˚¥úûpy¬1\®ÌºíÚ£¶x>§ªbÒ+†LqZ’"Ç§—j1Ÿ⁄<ñ„`£∑Æ&2@ã,ãœ£v¢ÓÖ}ãûﬂ'Ü∆◊y«°vhìWò∏ö†±=˙?ÖÀˇä‚8Ÿ≠¶L…nQ_î£˚?~ŒÄ™µ…¸◊‘W<3ÜÔìK“‰ë¢≥:4I¿ºPBÑ82+=ï>‹ı«ì¥£"ù$À∑√®*T4O‘§%‘â$~Uk[Teg∆	s•me0S˙r´ãKa∂¸ÂknØæ„Óú∞≠ÕÕ-«J’RŸJÏU)Ä…™#F}5d/YÇÈ|‡Amçx%+"öÕ\(Û{saI,êy7R¬≥KìV3¥Œ‰–©l¥¡}'îçú´ßÕ0HŸ¯MG≈}œ#û$ãW¨úìïWUv1¡aÏ§·§ù*∆£<å¢Ä0Â»ÔmgÒlù÷-`Ê_îÓîü¥ÛÖ^Õ‰Ñã@±¯úüÖÂÕõ±3…îñ∆…˚ÕnVOèçÕf„yæŒQ∆4„Yp?„~€`°‚&‘T”öáß÷t€ö
ÿ¬rÎ!Mø√;eô‘¥)Õ+Hcä›áÑ{∫∑Lã™≈7≈}–2úB?4ç€87≠!êâU	Ï0ÀÕŸ‹ØÉÃ€,ÿùf@Ï¸äÁ¥SÊ®fä87ïπPLghZwœñ0¨m≤<Ñ ≤ÏL M»äó‡U∑Ñ˘ç'6 ¥kEÒ%„K  ≠¿*Œ!O∫+Ûn,/€1÷fΩV9ﬁ‹πÓﬁ~ŒP∆upÒé˘)àÛŒCÕóEºÎ|X∏£∑S∆XÚ¥IÚgP#∑ÅhXû{ôHÒóûá=!J<ß_<øHè>~¸8é]rDIÊq”Ó>QÚ(üƒ ™I≤∂$!ƒ;π(E‚ 3Ù'™úa◊˜‘á≠xıaq[gäŸ{ dA3J‰M'7	–#.abÚjXå 7@~w?Éy€0Åù§øÊxbﬁ¥6	Öìë$%öàﬁÚf≤‘• Ñ0‰ôøX@8=Q˙â9l¢ÙSK¸$h-ﬂ O"oÂ£æ`Aƒæ√S˘{q<C∏ÀÉ∂â°s¥m˜S4ObnãY¡ù{¥Êq¨˛áÚçòÿΩaGÑu@®†∑@·~Hwe¡ŒåA…J{´zÎs"·î0\J¨>£"X[xÁjÒ)9ßˇFu,∆—¢§--9yÅuÏÛ¶°ıYk‚c9‹Ì	Ë ÑBË¢,Xæù©m¬ã3±ÎÜ¢îÌ´eBñÃg%æqË˙¸°´ÏO˝ì‰‘ûè33˝›Tå∫Fc¨= ¥íñ|Á¡im—7±Ò≠≠ª≤ ˆn$Cê≈‚— 6s˙Å+<F™è¶∏˘s;#‹Um„_Îß≈ñ–ı§Rê?Îs„◊397í¥Iª®ÕjΩÉ‹˝”PI
Ub˜]Ù¨ÃÇ…I†ûå/Ÿi|tπ[Z≈UG1∞¸î{öÌ∞fgÓMMˇ¿/∫©¥⁄É|E»“ å·œ¢4’ÙYbÒÖãÔ	=nô; [&[6ﬁÏ„ÕÒeÙ´◊î∫Û%ì∂ £¢·5G¿4*ÄúŒÂF˝ö±PîäEP–lù≠∑˘Ë¸r»cijÈ'÷/Q3=‘°˝Ÿ≈ñˆ¢E)∑‹Q…JÑ8kJ&áô´ì—õ§	ÔÜUÉMbúÒ‹jò%ynÅáXª∑≠qQZâŸq§Køx6®§—ë)i£˘œb`
˜6‚Ä§ˆ˘ ‡4ñIWsvw≤—i>h§∂B¡Ç£C*s¨àf<€E,?»!£„r¬|¨Ë˘XmŒ«,≤i§œÖënJ¸Rı3á√4ò¸ñ^Ò–.«√'ﬁ@œÑ2&¡zÚKg'÷eøPAﬂit™z6%ôÔÀåÚ`Ä“
∞u†’≥m>ﬁÔ™ºŒÈ∆’%ﬁé·;‘6ipMˇKŸ°4hk¯¬}5ø.∆Ieö¡¸§§C:§ßˇ ˇ4^\YZ%\nedÒö›¯	lìÀ«`o|ÃsVXa‰LUÀ‚’‚=Á¡¬Ÿ…‡„l6«ó™/OasÃâ vå9‚F∏∫≈‰UUVsfñ‹X	¢J„jæiΩg«KÿãÒÌ]!h‚åªp¢#PÇ/å›WFòOv∂D%C⁄6 †πQö˛-§6‰≥•∆¶0∂”èu€\‘˛gõäùÒi[6î√ã•¯∑Z◊ûùKÎ^L˚£«tY'†⁄?vt[;\√÷/;≥)X]‚÷40ÏıvO≠¿Á¿caÿ”RnÈ£µUÿ§ ;€Ã8os\ªZ\yL&˙üxé?	q&C T`pgÖ’o	q^I¬8E9Á8ú2Ò√èıÂ˝ó™ ‹Ë&Äπ/ú˘ª'Oü¨<ﬁ9Ürûz˛ûJ*ßtsÁ?eÑÖØ–=Ò{*Kà˙J*l–Ω "(©ùm!ªÑ'úë˛(¸’ÁMˆq§Å>]qA˙˛≠árVü
ISSô:f.2¥˙aíÿ)°—≈9!îj‰ËªóÙüıc'u$Ÿ»—w{Oˆ6ˆ‡≥à≤¬«!\Hãê2PY}3:%!m∏„˜ÍW7˝Uvï[íñ∑[ûv˚Ô…ôÉBÎπÇWÈEy‘|»8≤nGn2æÓˇH◊$†m–3∂¬£ŒÈàÛµÉxGNYÓ.;E7)±Ω	E2Ë⁄?2Ï·¨î[êE–BÍ–¸HÎgwµ#˜ìWv:Ëﬂ$˘vZ%”a¶,ˇ√b∏ ≥Ê$oØ⁄√«ô4^±*¸è«|Ê»È<oü˛aT^QâÓl:¶Ö˘p¸
„√ˇ¬∞¸)1,>ñ~÷|ãOa¬≤≥»P»<ê"C
Ï;≈˝ˇùŸdéùí•—Î]Œ›ú™⁄ÄÂ…Î1Â@NÜoD˘∏Wj≈´d{:Õç„MπÎ(éGñFœ	ƒÙ›Ë#4“øŒ∆ÒEù:??dvˆó»¸¸ÆåF÷«Z+#CœÛ*’ÊpïXwÑgıÆIú”ÅÊ…∑¬—ï–6˚î3è⁄3∫*Ëo wΩÍÎB=ŒÆÅ9–‰’˙E=,Í:Ô`ò£2 Eåyc5Õ/Dò≤ÍMı+Ö®Í€oÂŒ	úØ˜ÖÎf(˙cáÿ≥ˆ¨™π$ÀZ*iL´§õä∏ı≥GJπ»…ã&,_âÖUtˆ!òÿ qö1ˇz¥“€8û=€:cÆ¸•êì» W ⁄ô$c¨N›oÄ5VØ”πª˚´ˇ  ˇˇ Åm¿