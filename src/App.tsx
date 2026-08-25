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
  const [deliverySector, setDeliverySector] = useState<string>('');
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
      setDeliverySector(showRequestDetailModal.request.sector || '');
    } else {
      setAdminObservation('');
      setDeliverySector('');
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

  const handleApproveAndDeliverNewRequest = async (requestId: string, currentRequestItems: RequestItem[], targetSector?: string) => {
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
        if (finalDestinationSector === 'Farm√°cia') {
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
          sector: finalDestinationSector,
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
              sector: finalDestinationSector,
              location: batch.location || 'Almoxarifado',
              date: new Date().toISOString(),
              responsible: user?.displayName || user?.email,
              responsibleEmail: user?.email,
              exitReason: 'consumo',
              batch_number: batch.batch_number,
              expiry_date: batch.expiry_date
            });

            if (finalDestinationSector === 'Farm√°cia' && batch.location !== 'Farm√°cia') {
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
          sector: finalDestinationSector,
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

  const handleDeliverRequest = async (requestId: string, requestItems: RequestItem[], targetSector?: string) => {
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
        if (finalDestinationSector === 'Farm√°cia') {
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
          sector: finalDestinationSector,
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
              sector: finalDestinationSector,
              location: batch.location || 'Almoxarifado',
              date: new Date().toISOString(),
              responsible: user?.displayName || user?.email,
              responsibleEmail: user?.email,
              exitReason: 'consumo',
              batch_number: batch.batch_number,
              expiry_date: batch.expiry_date
            });

            if (finalDestinationSector === 'Farm√°cia' && batch.location !== 'Farm√°cia') {
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
                  <div className="p-6 border-b border-slate-10xúÏ}Ms€Hñ‡}~Eñ ;§∫Mä¢$hdyUíÏvè-©$π¶7"R⁄ ¿@}¥J±á=Ìiˆ¥∑⁄>tLGÙ©b.s’?ô_≤Ôe&ÄêôHPî-†ªdÚÛΩóÔ˚ıÊ÷ˇÅ(Æµì%2ù8ﬁqFÙ…\BœìŒëSr…wª˚NB;è{Ωπıﬂyqr˝k‰B‚Rr˙ﬁ¿KúÎø^ˇçÒŒ=˝	|ˇK∏∂p≤§Èx\È˜<ñªZÈı¯ F‘ı&£πıçA8;¡	%!qﬁíê‡∑±—çàÎƒ$Æ«ÕáC‹êƒtˇ%a‘][´∆∂∂‡zßÎˇ†˙~êG=Ó,ìx¥:Ó< Òÿ–ŒEgI≥Œó˝”Ñ∆I¨¸ïêÓ±Á'4jG‰…:â∫1¿…ì'O`¨>|°Óøıèˇ?{Ò>M&QÄﬂæã∫.<Ol$Û∫÷GŒ∏#¿∆/5œ2É8!ÿãÑébÚÑ8æøœáÕÓdcÙÿ ΩÆò”;œe#ÖØ]œùˇ'mt[˚ÄXb√ÔÑ|†O.yWWÊ'•ù:fPµ†Gé¬»x·ˇà_˙=ˆ”IxJ£UÒã3:Çøâ¬I‡R∑”?˜I∂Ì+$âú ˆ/:∞XÈÀCÒ"t∂@Ï”sˆß3}|ë}é¬3¸Ï·¬v  0û?N‚ƒ;æË—‰å“ÄËjŒ0G5¨÷RZÑF˚|4ãP’7êœ‰,r∆§0vkø∂EhF-ö" ÇC`=œN†32>ÔÙª+d|—YÃ6ƒÍv”b0Ñ|/@©Ÿ†ÌŒ |7	_è«4⁄*ÿûø™ü–Œhäâõh^LGﬁQËªVs¯Ø˝‰2†gd⁄@4Ô∫'Ú28>≈€I‰√vkút~ÿo›‚¥.ﬂ≥˘ºYÏçœﬂ ßlﬁRqÎé'Ä2\Îû8ÄKÉ0»Œôóo'πßßW˘Ösé'ôƒåµ∂∂⁄}˘zsc˜›ﬁˆŒ÷ˆŒ·vã<%-	1˘™Ûo{)E»æEVo‘Ô∆ﬁ˛ÓO[iøtD#«w≥û”ÔRﬂÈ≠ÙæøΩ˘˙@Í=
cöuÕæH˝≤Ô÷ù∂2R∫ÿÎ…∞+µò°`´∂¡´˜W6 ~iπ√˘õπÿÓü¯≤{„V˚"æ4ouv»Àñïü ˚‘âC∆PòéÁ¨%1+ =P/…qBŒø Ë,<ÍIÙºˇ§dadGıy&ﬁi∏ZùE˝Ÿ/Q-©‘≥ãÖßÃG)ûûã∏â≈±Ã˜ç≥d»ÿ·9åô˝∆1
_cLî‚L ∑PwÊ¶g≤Êï)n}|‚∏·∞O±ÂÊÚ·ç£–ùíwåÓ
üDa0¨8#„sÎm˛ﬁü&Nêx…≈;¡¢R˜jvúΩn”ø5tÃ<áG1çN‰&Ì–JÜYXõl√ê«N-€K‡gU[∏h…?ÕïG~e‚W”%Q ]ÚU≥¨6ã∫v4IX…öû¬`ñÁ√ìÀˆ<‚HÜ'·ô{∂h‚x˛´–u¸ˆ%¿px∂
l Ñﬁ'ÆVÒπ™Öï(Ú8≈íåπ]F4ÎK§6ßäëôêëcWI»oÅx‹°Åã¢˚,∏u÷%óÇ)êy◊Í∂Î'h◊ ?°:ó_k|WLäÕ’˝¨ï/qÃ≤˜åÑÏÆOÉar¬^Ïô0µ™/XÏÛ€PB’L@3†ﬁ⁄~L ›ú…-èŒ;Œ$	Âó ÉIÏ˝XÚÂﬁY049Æ cŸÈèJLﬁ‹˙N&#GVµ¸%,kZhpÍ9Æ£S≤®zV∞πı◊âÁ√<H$*¡∂ùÅ›œ•ößH÷8Õ°6»¥z@Ó·∑’a∞:˛p2"#h:Úü8–ñ?
œù»;î1÷ ïJÏ◊<Øº]‹Kæîà—ì£CÁàsíbﬂ≈I8¯–RC^Ê≤É:•)KäﬂÁ<YN;êÃ˚˘ƒs](a≤
·“VèlîiTê¿¸1∂g÷êõ¿ˆ‡6ìˆeóÒ‹üùÚ €ÜM°8:ér`Nú»Ö≈#2\«H⁄^0'±wJ…)û∆ÛBŸ√€·Ì®%€¸ÿ”	Ÿ\%8òD,æ†|‰	◊˛§ïiµh˛ùóSOÑZÌÉmØÎáŒ)»Úœ?ìvÈCãgN4∫˛e‡9'JmHøŒœ{˜2∂ë¨ìûÊ18~4øx«§]]©¬a1o–æ⁄*Gg|û∞v˜ «ú!ùŸâ¬ùÍTa`ÖàMsƒ6jÎuÍìaãGrúÇ≈«‘q«/"±„√Ô˘Óì0∆QΩqÑ°ÅùAöx"ãSœCB·$5«ﬁx	ËQ¿ü,!÷È˛È˘ÿã®ãÿoÛœÏ]ÉNüøP'b/\∞ww≤ØÈÎ¡&ê…≥ñFÅÀ˜‚¥”ÌﬂPç_Ø»LøV~*ËÙV*äº¢É›]Bı^m√´Ú>Z#≥t,Wï®*cäÌHZJìçç≠∆¨tºzoí’¶3®†ÊÊ.õTä‘ñ1P[òZ≈ZY»^ÅìïCÄè[–Ê'W¡‰SUﬁ»‰Sû…G’BÁ)û?ù’˙Âù´ˆ}XÏïçù)9–(_Õ6≠µ—X≤©l¯4JÅYÜ>úƒb9	ÚgégπErv´-Yè$≈viΩ3™w◊‹Ò=êV(ŸãÆ=yz∆K?Û’.XÃrÉZu≈”_“5o¥‚ñ+ª="¿Á›A3Ccâ∂vd?2ôàÅ
‚Bﬁ^’:%Z.ù5©lu%»Ô$íw#Íƒìà¢d◊Çn´·it‰$Éìw¡ëÃ˙DZgÛóab=çá˘4‰”¡Æ’àBxŸú¨v<\ºC>ìfÀ˜æCôˆ@2PÉë8Æ£—igû¢~uˆóÔÛÈﬂì™V~∫<`Übq'[ù⁄qsÑÚ†Õæ≥\≈z)Ò÷Lıé
 hnezΩú¯M⁄ol¸R3òÁÆ!Ê‹o≤y‚E¸hÛjä≠‚µÙ´Õ´#Á¸«©^LU;? äΩ»ßZÔD÷ †Ê≠KãœÌƒm	¬¬`{E˘ÑûR2â√V”∂wscò‘¡˚W©ûPËQ≈poZîyè√¥⁄≥Ú—UçY‰™∂ù+{˚÷Â{+C?◊ol∂“1H–µÖ„êµƒ]êπ‰ÇÌ0Á∏sKûÖxõ
∏ú9,4ò3ïRãu€cîiÎÕwíMIàÀåG‹íT‘íä˙ö¯‘ø\Õ´=ın≈C}oË˘Ùùü`u<’∂òÌ âË–â	Uﬂ˚≥É^ÿé0«‹éÌçèÿM{F≠Ïë«:>rºÛê+fcÒÇ	˝-Õå.ü±Ô∑ÏY∂Ωs∏ø˝¸ıvÎõ#∏È…èÔ˛Ö˘|	ûﬁÇL}9ﬂÃ—õAÃ∞j•∏ Fî/}t¥*s†•ˆÍP´ÓV…˚Õ˝@îù’≥π≥⁄åª„§i#Óq˙Î¯u‹ÓÑƒçr%◊FÒ√œ?ì^Ω$êìıë„∞Ãïù1åß–`ß‹∑MÎ4=ø8jX=⁄»∫ó_ïà+Ø÷1¸ Wsù¡¨ﬂK%
•+˝2|ıΩÄví‹§ÍWﬂ§#.jp∫ZÚ -ün”Ÿ?èÎΩ˜˘eÉ¶xŸ©V’ÆøÌÀl•Ø.’p{E\/wuºOÒ≤%‚i$øÍ@k$ÈëO¢ô‚§·à©§R¶Ø@æ8´w‚y3%]rs7#\ºmnL˜D„äì†ﬁﬂ1»+83ej5Œ˙z¿µÄùår⁄Y(ufß,´K£µk"Wf˚a˜b™l¯ZE-£ÊZxèPqÁﬁÙﬁ>ÌÚZ6Ï ªÀ`P†?Ë,fÚpI/Ùì0Ã⁄1LZâìÏ0t‚§=w∫aú{&π }9©<L˛x˝úuë3"Iò8>Jß¢/: vÁÓì9/8ÁÄß≈2’/ÅJ.-…îZ‹Íkg•S≠ÃP√:WP6ç `‚óJ€®‘≈∆#ÖŒU£ùΩYA¶kzrÔÀ	$∞”ÿ|ã*‡Q)ÉÖ^ˇ˝î˙péhÄœŸD§˘‡‡†/vÆY¶JÑÆ˛‰˙Ô®Llœ•€57œQΩ»î#˚o_ˇ≠„s}ßp˝Ñ1≈gF!"s∑Ù®Ùÿ⁄¬F‡°˜˙^DcXcö>Ô¬˘úq¢óø!å¬≈‰7ÈÎóHÎ6\ó˝PÑ÷äÌùkÊ∞Ë¶‚a4à˘¬øné·Œ$Bò¯3)W5N©éL|w@y-üjë1&#Òˇ…%êfT1¨í^˜Ò}B0˛•W&ª_ÈÖEÈÒ≈Ú„*’˛˜/^ßs÷YÚõR‚7ÀÃ≠∞Uí±¯É'ù7è{ß'os5ˇC<ô¢˜[©©(´k8:Í<™Ä55|)å h©(ÎÔﬁ|ø∏π¯xÒ·[n@-èKs˙≠÷˝+êÙÕ˜=\|∞˘∂®ÔﬂDK`%]ˇßüxc_‚vò	‡î˛YÅlJÏ– MÍø∞€« ıTÅõÚ€œœ”7ﬂ?[Åˇ-ø-zIáÎ Ù√(.üïäu˙É ≥˝ÂjÑÓõÔ7mÙoøù´R`ı…®.◊ÄGf‹y…ìÀ'p}
sGfÏJ•EÆBáÕp4Çµ}ÊQﬂç	û|{ë”ä¨ø™9÷k˛˜Ò‚Qé*K}Dï¢∂˜Õ˜€∑W∂aiá˜¯3Ó,íëªöÌ(}]fjªä≥Ø<xﬁË†Z@£ﬂ[Ûù#8è‰)¯°§ÕÆ"H
ÿw3ê;/·`~Ç¥:†.ZÃX' ÓΩ`<—Ëyë3A[∂ÚGŒ-ÿDÑÔ°i2B◊∞I8‰Wr1ê6ßÓms§'pö”Ë…‹N≤Ø¢ƒ!Ê†nUZ6A$SÜuINûÔπƒœáÉIºääb¥j‰_2rƒíÓ§¸Üz,ßé?‹:ö¯êt]t”≈Q≥·@% 9ÜT‚áÙÕˆe∑€Õ⁄πO“ÜV	Ì&N4§IóuV22(œ{7•aÓ40zclâdÑ¥7πõµÁÃõ@[uûôMåˇÿ°g¢ıã5ÙNæ™Är‘ˆ√ﬁ5‡ø$Ñ1<UÌ˛”! øJx|Í ,õûh‰Îé”1âÆUÃŸ)æ‹.¢âQëä<—3úØˆì<RØ:Â€»ü3-C#ı*™}J÷M"o‘ûØW˙¿r•ËC„6»+ßÿﬂF9›„(a√Ëh—~á?ﬂ'ÍÆﬁŒœ◊™ LT-çïX’¥_Ø‡PÏºÖ≤EceD|ô˘m$’-eE`UA’(õc§VåÍâ=ßWèÃ≤tç∂√$X|Ò◊=rEÜ	[Ó9´úÍ6!>ŒˆäDÀd˛1A¿Â CnfdÅØıy|÷¬1„ØôUﬁ∏Jßçü◊Ÿﬂµ˛å©oSÊöµæ˚üÇòÍ0µ™∆√A%å≈‘÷ja≥*Õ›üô⁄ËN≥éªë7§#£h¢KwZ¬ 4zö‘dho¶LÅÍ É5!“8`
:¬:«U¨Édn}ì}JB3)Ω€9sÎ{hãKB≤ç_5‡Ü3¡Ã≠oÖé»¨]O°‘Í%Ë£ÓÇª_ÇÄ‹@që™Cñ-¥
µö95Üù,Wî>ÒHØÅ”#÷‹:NŒÚXÄZπeeè&ä^OÀ3*Ó∏.¬.ÆÎ~x¶wc¬9C≤Ã<¢èπÙî\‘∏/Ê4^~∏hSMT 2°ÁÆÛÆ'\èw'I28S¬±éÓ€RÎLeÃm5jhJú#ü*»cÍWD—¢ë—F$˝Ñ ΩÕh-9°é´=≠í®≤«>=NL2vrRP™fªl
	Õ‘†¶s≈:g–¿#÷¬»]7˙L?∑ûjØpÉ÷í˝|%∑ÈΩ¢∞Ÿ<π.‡#LÙ¨≥Ù õÌ"ü\™âIcèÃ4qÑe‘}>Í◊7] €££ÆyÃS,≈ytÅ´áﬁFëT,<Ã=¡ìÃ»ûÙ∑KÑI‰Üü'T≠|/|ûx’8Ù4ü÷´Îø∑0t;êË√¡ä·πül ÀΩπı4Jˆ”≠¬#‰ßËı_CÚ:∏4∞¨Ö©s¯-“ÜSb-9
›]£≠ı2DÊ~uüx ∫ÁÛfAO mí€!ÄˇÿË¿Ωñ∏•ÂÏ◊$f°ΩVÖåó—àì_‹úÉ3´7Â‰óùé/•ugÃ9˚∫wï"ÿá≈Ç}T∞∆sIÿV¢W—µÿ'=b•#¬K»fy
úöÁK‚Ÿdåë)ó€N≤I€j›'4⁄Fç5^(£¸8Œö`ærB}¿Tdó#2bl€{ÖÃRjf&º^e˛sîdÈèÚÄ')AJˆ›æ>í'ó¬πﬂƒõ˛#§RG≤Q/nUZU'&y§¶¢5 ÿ©)Ç≠âWÌ]œ/ﬂ¥FCÄ®÷h¿˛y˝ˇÚ;>˛˝o≠∑å42‡∞Lû…zVœäÏ¿ŒÀﬂ^Û'_ç=÷”ÀßY~ ÊÑù;d€Ñçè+ø§éÿñùSÆ¸Úw“€]∏Òøx…IªEZ∂û‚%OÈ∑Äß§e?≥È£¶'‹Ë&]’Cy=Ì€µù@≠€lzOÛ≈<∆‡∞§/G·Áä~K2˝HáŒ√<T–Ç«•¨fEß⁄å@ÿaÖmºÃo`ßçﬂlz’'∑∑äkaè}>dz+åù!5£“ËgÀ)5HXÖèúdØ‰{]˘~˙Ö”q7åÈ≠“qÏ‡Ø{˘†„∏—_ó‘∆ınûc•L Â'Ó1∑G—OBÃ-2vµ∂êh«xÕJT1…ãV“ÖÖø,Âˆè)√zßïêã2r.iæ√3Å«÷’Éì≠ÿ\jæ©mÉ[%c£îL¶€Ì⁄∏K(õŸ€›?‹ÿ±AññóÁ÷ÂoS7ππ˚joˇ≈´[ªsÎ˘Á©õ€xµ∑˚rcnùˇ;u3ª/__ˇØÎˇ	É >ﬁº1r∞∑øÒ?§&˘çl«´ç≠‹¸wÍfûÔB#¯˜ª¯Ú˙Ôø`{»?Ÿ6UÔz√ü2”P36F+]¶Y¥$äüì¬∞ﬁ¥R€U˚ıŒ<2¯{Œ N£Ω∑y»æo:ﬁπC⁄õ`ﬂûEN<I˚Ÿ>˚∫1á>¸
ÿ æˇ‡≈Å3Ñ?<g7∂ÉSÍá@Å€€;?±;œÉ_€œ_≤Ø˚°_ˆ_äæ#Ëò7˝ä¢∫˝ä}˘q‚·cˇÃ}È±üD49 å∞≤%.ç€Í\≠-õ$OâÊÌ¬"Ÿ¥¥JZÃå^õØô"◊6/l?»#≈C©ûFAÇΩâÃæùÙ†=(ÂÖƒÕ≤	Væ"‘ôrñ˝¬§l:Æ€´Ê,,R\§r2~≥¶º•√DB\8R§oS6òc>Ü∂•üßlL"sÎ“ó)õì	0“∑),PÆπı¬◊)õ,“>9îøOŸ®D>·»œøLŸ\F~Á÷≥èS√'ﬁ{¸”îeîn=˚8eS˘πÅut“œS6ñ;Ë°&>NªÏxh¡í„?S6¡ËÛ‹:w÷jªﬁÎ Ål`9$;∂ÌÚªØàA∞2&Zô≈õÿÆK÷ÎÌxbﬁ±á.-òÉÓ»ÒQπ[ÊƒØ*gä∂ŒBÍ0]\∆a.IŸ"ÅQ≠„A≠8Ÿfr{uWf'¥óxÇÜ{ç’˚F
§õ*~fÍ´¡!óó®€<‰ìπ≈Fﬁ?&µ≤KçÄ%)MÙˆUÖ4,º„Ûñ“dv£¢ÜıÒJO6À‰)fd4K◊ò—éÄFû‰·]öéi3ÎawÂØ≥rı(∑[AŒˆc°w)åfÏD1}$Âx«õ∏á‹zwjâ]–aÒÜ®êñ2”V±@Ç¥ç
€∞hâ<Œ¥†À§{˘óYÅº‹ÊgÓ‹ôtP>•˚∫±~q>sÖ83Ç.πÕô˙–M5ÀuI•ıô\ΩÅ©˚ãWœˇπ^å¡nÍ˙øÛÚZA	}«Kå‘ÛÅz†\‰÷÷èg¨ Ê¥*R11n∆KuUf«KçŒúV≈XjJù†e>£Œ8Ù,gl¡.º¡	|8
œm›∞Y"'yƒm
&√»π`ßi)áîM„8êõÉv.O›Öº›bL÷uãÚVá¶≤#Ç∆üinΩPÃ63∫6÷V~®∆Ê?Â—\Õ'Uz≠¡—l«Y∆	kŸÌ5ì≥z˜{µ$Á≥:∆s&í	Ú„»–"ôﬂü’ü∑h«>>ÛCÁ.0êS±SÌ¥LÓ›‘1Ø‰íÁN∆>∫[Tv»ÇÑï“ﬁÒÚ@∞+,cΩ‰^≈Ó Z±;+EÖXm¥mi÷^‚C«[|¸œ¬2Ö≠o¡e[PﬂÕp|aóØAº`ÈUâf©^◊…¢]ÈFk»hÓ¥YÇéàé`À§Ó–°ÜV!≠ö:j7ÉèBˆŸjè;üºµ√»âO˙M¿¬0jıÜ«Æ>2OÔÔË‚Ú‡'‰¬-S. Ytc’àÁ*¬’T]\*R[ÂÒ¨Å«\ÇÁÁÁŸ-ûw_Á/]Hï√ÎH§ÃÿÂN™Vf(ÊR7MeíEÂT·4¸”ÑàÁ˘cÀπƒŸD‚ôÃBùl@õîoK‡d≈ÂíJó*-ÜûÇ’Q¨)S¥jT=@ÜlπP∆KÀY+“∫÷%’^l:¡Ä˙N§ÿmˆá∫ÂäYŒV’rïg˝¶ˇVö∑.°ôz=§Ë?ÓØÙ+ê*ÒÅ\ãO9∑⁄∑HÂå–¥d∑äkŒiZ√ºœRèìg Á`≈¡(∂«\ö√ˆäß?±ﬁQ[[¿ÑπÚ=uñÌ¬´y.yñ≤r?G/ÇStbé.fù<˚ÕÉﬁ€/)}ˆ»-T ƒ¡yÊÏ'«ÆfV•¡.`á}iºµΩà)g™µ5£<'%{¿~Âå,ßŸ⁄{aî8eO9UÇÌîp®âf˙RÚ©O1-ÛmR∆ÔÍˇP@Œ X≠≥Rkí-)6HÕv‹Rr22ÚQÜÀ´ê:‰¿Ò≠˛§<èR÷Íæ!ÌÏÂ˛ÓÓ´^w	vÃ¿¿xÅLh@«e*ÄDîØ@@a}k9TU≈Á%}ŸñŸâõ’J'»ΩÀXÍîQ¬º>öÚ∞ ¯Ò[K‘8ìéÈ‚ŒöR¸â[-]3mä
˝ öÿcfõ≥Ïì¿¸3œá_#í•ré≠°ûó7X~T.mêW%aôî©Ûı)’X”$õ¶X±<âf≠˛fâÚÜŸ&‡äj¯2£®ÕûR£·¥’¿îèÀ‰$gq’z˜LC#)Öé∆†ËÃ2 Êäsß0XRÉl€¿ù®+∫ˆoÊì2(cödı–Yr◊∑F∑f+Oj}ﬂö%ïÑ∏ÚP?\;Û@¥øi}™µ:mç⁄äP¥ºº‹[⁄‘)fı∆ÉIa‘O-ë69õrG_çpiæÕÎÑlüèÅ)-¨Ÿﬁ÷≥òc‘0p∞#¸åKuü®ë@Ìu€¢zKâüX$NÈÍÃ%‚T™(≈œ)’[@÷Y®ÏÕﬁ“≠ÊÚÓ!Æç3»™‰uÒÊ∑zQÅ`F/N[*J!∆AZ)˝™wèÖˇ2[Lª@É]küΩòÒáB√ÿÒ¸ÓÅÉ’‘JÆ«UŸµ¯≥≤4ë4ey¢Ã¢ %*ì„´€ceAuµG*±jµ≥á≠´cÈG¬Söô –)Î‰ï*9-¬»\|Àã«a jıç“Bx´Ü≈ jß^aæ!ñK7V›4¥∂}¡Éf;a©≤TÃ“πæÚpeiy˚≠Hdû™	 ˛yfY£óEmHm≠èmMÍs≥ÛÅe sΩ\ê#ƒÈé3~·ZÛˇœÃﬂ¥≠	c0}i„˝…dƒ˝l‚˛/%„ì(˝kÃÌ+€h≤Ãì≤#|Ø	 1ò¥—ytïT}.Y,«Œ¬FÎjût»e£∫¶1÷NZÔbaér“c§≤HÅ÷¥iÖy%Í"˙™°Í“=a.ã∆rÚÑ›FT≠wıP∞s<´û“‘˜ÌL˙UôF"Ω?&ÌWNr†|ﬁ%)¸@¥hr¨)ó≈>9¡?2{åI‘XÖ’EŸ¶œoa^µRxTZI∂ûu+q+πã∑Å9ê•0+á=‡w¥`_Sú√JáSÁüV˝#»HqÁö–X8Ëºıë–Ù,!CU&Ù¨≥‘ÁpÃNFKCöˇ«ÜX§)·BÏæ–¶Ed“ÇﬁuXØ#˘HZB ﬂR~ı8˘€©pR{È™*‹¯<“sÇrFuoíˆ
9†#'pbSAƒ:™bCS
¨õé?ò¯ FÖÜ‡◊è¡E÷˚™|T%§xÂ	¨¸º°Œ#!Ì˙ÃÉ<´ÁËà^¢‚«˘øc/pSÔ#œ-‘yÁ´Q]ã*llßNmÕGuFÈˇb´q<_∂œΩoƒo2Ê-£ÃÊ$M&Q 7∏Nz∞tå®Á∑•ü~CVÊU¡c”È°≥Œ[ÜtáWÛm]Óäb„vL˝Yóæ‹ní¶9T:OØ_2õ≥`qIdáµuÄ˛kŸ¿-ÍùSîüè`ÏÉêª∑Nb C2∏˛±7d0G◊ø$<‡fÈ’¸æFt÷V+ﬂS“ˆÊ˘ÙÈ[°™Ø¬ƒ;Ö5pÑz©◊^1≈«£J≈ËZ~}∂‹Äâ4 <3@ù8⁄-§ìQhN˛#…‘X◊*å⁄€õáª˚oz&+ó!)O…‡øT)¬doßŸl∏.ù—46ˇB°?{ø8¬¥ˆ~7≤…:≠úâ«ùÅS7Ï@cL1ıiÄEÃ‚Œ√JZ∆Ó3ñS<∑!¥l—x‡D	]¯	ﬂfˆß⁄∂YH]⁄Ì›É∑t!Ó<¿˝ƒ˙ô¬€òFÆ”⁄ˆïÖçS'ÚÍﬂºÎ`∆ÁÁÅL^Ûœ^4◊L¿t‰laLG‡$◊øDc?S2èª°[µ∂ñ¶ÄßãπâÛÚéΩAﬁEÎÄÏ£o0I®IÈd‰á‘ñ
oPU~«Ωd⁄ΩBÕ2i´%À`/Ö£bö\6≈#≥P ∆`alï‡∆6?HQ)≥ΩK:d≠–!˜˜∆te¢\Lv]ò}Ë_ˇ:®)u7T7¸
ú¥§/Ï
/KÄf´FSÆ…WVï˚∑J]⁄ZÒ2ñÁ6ÇÚßÙ}:†c ic°6≥FQÊìB</'Á§fËIL†6%ﬂQŒç4Tôë˘h0Zé∞dÆäX[¯aˆÄæ€«k`f†nó±bˆ©ÀÁ~ ∆nënÚ„Bw ßZfãê	˘ÜÎF4éÊÑ(ëf—H»ØMÌPGÉ?nÓÏ˝˛3Ñ€+JÖˇ-¿ΩŒî©OÓ"Ù‚é‹t±Öè∑wéÒÿs∆–‡°7:äÄ≈&Ì›1ÛöÒÅß˚˝ﬁˆÛÖΩùÁVLàU Ûq§öûp≈/≈‚‘¸µSø¥3è◊âOÚËõ\ó¿^»‰%π›ƒ-£á§¡ø˚ Ù§i*ò:Ÿrtîaù”'©RÖå§ Cáı›Å∏Í”ñkIø9Cä¢ÒFÕ†s%ﬁ	…8sπL·”"ﬂ±mñ,õ Xç“˘{>ù´ÿ‡ÒÒdŒ√â.¸∆‚igN<◊•Å≈;9ÕkSÎ¬}‹`çÛê´c‡˜¯i˜MÔ≠Mô84é„∂uıxüÄ‡„ÇIBû¡Î˚ÏÜmç@˛z7eh‡BCÕ ¬ô≤ìm—&∞5?!NL‚è0À]58˛≥o9âÛzˇ%_>õ˜Îë°∂_y∞bq*Ëlì(˛£§∆Rlk ÃL”ƒXÏ` Ó,sl„ü{
“œüC-â‚É@Fµ‹M9›ab4É4X≈åGfñ´&´{Öm6äRÂ}@!xπ”,Ü˘Õ©‚yu ≥¿A\ÎOÖ—∞n„h§<√+ã”¡á£a/¢ß=kv4d”gˇÑG§É /HØqj©@ÕV ·Œ1ºHeQau˝˝0ƒπS/æ˛KxÁ5m€Á´§øº–[~≤ˇ‡≥”≠•R[o¯w´6wØ^nì€¯⁄6√èÃ©j:sê≠iF ¶7.#i{
Œì	«&TåµÉµ–ç…$óÑ{ áQ’%⁄%ˇdÌñ^Shõrƒi˛4Î Û…uúö¨n\xâÆ|‰aãÄÒÉpíÇé∑k
™´Ÿs›küE‰óÓ6—âœD>UÙÓ≥0†∫‹ÄjV.øÊD£Î_j‚.a‰jÇÀ±È¥à8nz< [èFQg¸ï<= ˇ∫û~∏ΩË/uêó÷¶9‘≤|⁄êÓ¸íÇªÛÄÓ˚˛‘K>RàwÿÕ^3º§0î˜’ò3Î”πıﬂ≥hVP˘‘A'·&Ñ6•–HdÕ´ÿ≠ÇË(óÃÆò	¨Ãÿg%º‹¿CÓºèœ–		¬4Ê-0ç¸Öêå2GaD$úáù#¯¨exÕG“h√∂buûì¥:¨  `¨ÌZ∫K%ˇ¶&Ñ∑‡eIyuÆÒ Ï˘ïßÜ(º’–≥dzÃ`	 π_ΩÖ”¯ÂëÄ•@B‹>∫O∞¿πQKT+—Fé±‡íwû>hDƒNËi>[2dåÉ≤H^µ\⁄_Xkı≠™ÿÎ4Û.¨ÁsÎ,h˘)ãean!GC‰√ ßCñG"¨·Ù™˝(¢1Ñ
zn]ä•~™	¶&?ß	≥”0nS7+‰¨v∫∞&ö∞nõ¸‚†–≥í ºcj-é1zﬂÏ≈7ÄﬂÊÎ¸Ñ»q©˙Á0bπæ†o¸˝v÷RÌ[µ:–¬—àú¿≤⁄(U%’∆Q*r
±mãuˆ÷:çì.éô_v∫∫ŸWﬁ∞´nˇDÅ{è?Æ€quß¨{÷¨∏•¸≥¥ï„ÆçÆÆ„[ƒ6…ßG£Œ:ã a‚L3k7
‹ñƒUe’ˆÃ‡Uß∑¸´LıÔ˛ˆ±’^∫ u~Y€[º|oW+“N∂ﬂcÀûXÓIŒ‰÷º <ó–7Vk|†Vp‰ÎÑﬂÙ™5°¸¡æ¸ÜÕäﬂƒÈF“Wj·f«˘˙ ÇuN˜ŸıŸMª¿I√hO@ﬁ∂…•xj®kX|…Ú5xÒÄ:—‡D~◊9äC‰ƒß«	Ê+	«ù≈Ö>·GÀnq¡n◊`Æë©‘“I#Ωü·	W˚§d™o∂`"⁄£Òü&ÜòãåX”E˚®?GÌà—	*ó€Ò8'âÔ¥r¬ﬂÿ¨ÑWQEçªH#mvÙ¸ö ;&Wjßï˘?[ˇ9◊.˝|9cãt0∆¥•WÌyãì˝gz±ûçfÀÛ4†\{∑çánıµﬂïVÔâ3(´Ç’≥ı˛°›qD1Ï=v&~bÎ≥∞@x&Iíú XE#´◊ÑÉ¶¢Dß¡$B&ÙP$£ˆ·Äãìv∑Y{ºòª<?œZÌ¢ä⁄‡c≥õÃ«p‚—)ÌÛ´åCŸ^>!}∂ÎeòµrÚ©W‚ÁóAùøbÊSÈıÛK{jÙH‰O_<=xnWÃí±™$1œ Ã‹ÍrËﬂ‹’Á≤ù™)7(ïﬂœ0”âŸ	˛Ê≠.Yã|e-âlâMú˙·ËH?z„Ωø'ÁAºß…Äÿ∫zoÎ6«™†÷¢Ã5Øùv8üÁ≠/?RsayAõïä√(i∑ù˚‰àYáM∂åó„”ÕpÑo˚H;≥j¯}⁄XÏY=oõ‘2øÍBƒãW!	¶Â;Mã‚’XƒN/§≈‚$äC8@èıL;œŒ±¥ˆ°Ì˘ƒ/ñQ*tlÿéë.ÑÁè◊ø†ï
˛¡l¨¿OÃÏ›ñìé¢0≤=G¯≈·⁄˛€mêe˘\˝pü\±8´È ‹')±`ÈÆÕ*]î˘#+Fz:à_4û_&K „]ëÏ◊gÛO	˛QN˘”°Édµ†Ÿ™+Xeá6æıx’áFû6õv eΩ◊%J^kŒ—˜£7Ìà[aå<’„¢”ÈXXî§ÿ¯¸gÖvæw4p≠‚DãÍZ‡≤8ûß^î≥u3k∞.vW\{OZF™ﬁÄﬁ4À4\öéUúEπ/n’~á•Ì*ÂJñ÷ÔÚ=_¿GEu∆=Ã3$à€¡E¢æß§UcæE°≥g4Y;’åò®T®ŸŒ|2JwVYîÛ],∑0ﬂM¬óåc¿üœ“'ùˆ[V{≥≈j∏;÷ç7@"€¢π¨ˇzßnÑuﬁ
Rg7[Ø¶.õ
J”ô„gÆ·ÎÀ>$ªv—≠uﬁé¸∫©ÛMQ•4mR{~5LmœØ)Mã»uí¢í®ûùA6ˆ;‡Q≠πAæ◊ñÂ 5`s≠∏3€y‘ì†z.v∂ZÆ¿äÌúö_ùéQÆ·ok’‘&Z‹i+∑Y~Ò*
µÛ,i!xiÔR™›îBïL„T“Lé´DS±x5“À»Ω,Ï¥ÆH[¡ÇÕ+8ãß‰}á%Ñ[gJˆ·=2=¥ƒ¢¶Dz’ÓÄπ;≥r˙åÒ@6∏7Î”å’‰,≥∫]È∏ÙU[€n	´≠7,p◊∞¶∫ö‘ã˙kÂ‰ÿÌK¶·X%¨
€}÷◊jVJÈJ_ÃM[9u∑E¢,;˚y’€´+ÊÓz±s‰≥äõ5Ö¨ŒΩ§Uµ≠‡o=Ûj^æØ,ß¢^qeãöçtê´BëébÁ=˚˙aíº˙@.Uêﬁ|(‰Ø4Æ©PÍe†Lf®‹ø08ˆ¢ë’◊S/©8oﬁ∑Ô]~∏öWå‡#WóﬂÍò–[)¥˜≠∂<üêM5ΩÂbºôu°¯b‰µ®ø}>'^D^¬’J†S.¶g¨#ß¨{∆B‚D(`«¿¡h∏Jêãâ£8"dË1íqÊ#+A~
◊#„–ÅtÉ…?u|#4N@∏˛5bâÎi÷DwK3)©ÁäøZ˚“Ã*|fÅDÛÈVÁÃa
ÿΩ≠»í‘k°Øêi+ë¡/ù‡˙Ø¸¡˝`{X}-ç≈E≤¬‚h’—%ä»€¬ıq%:oóÖ2ç¨jS|”(È<™ÇÜÓ4VñHjëÒa._ÜÕ˝8LJıà”0(⁄ì»r9xNæ“°Ÿ.;RŸéÓw˜Iv.ô|‰◊1%ñ†˘ e¨£πq5yªzgP&wã&éÁ´ê´>∂ﬂ<^9={K¸·™†òâü‚xÙèÒe,Ê°.ú˚h•Y·\a1YÂ‘´⁄∫¶‚#–Ûc¥™`)§ÍZè5·˙( ˘XòÎª+‡¥‘z◊ß,˙*àÒpO†»éwj^X%˙Œ∏v°eÓh/∫˛ı‹Ö1qB"E†jzkm¿[úˇÿ°«ò|€!ˇıØˇáÁv|<¬6P©L\eèÿ]eji©—üYÆﬂ£≠Ç˚’úC:¬Qx,‡;fQ™Ãª+Ç˚„∏(˜˙oÖ{ÿúìT≥ß+LìJE¬ÙQdzéåÿÀµ›_Üg¨r“ﬁ÷≥Ås·s##À•3å◊É°tí∞ë„(uTd1:~Í9QsrÕ	∏Ùt~3}ÏQÒPUÖ=T‘UÑVOÆ“yû•£q‰çÄﬂGc‰eÎƒAz3∫˛;VeäÅe,‡–T•‹π'1QÙ8µÉé’…0Z£„…I…}æv´‰Õ[%#TJâ’/W \î“a1í^⁄û"Z.‚
í”ÕÑÒzF3ï>™òV!∆ya¶ÉÖ≈su°ƒ¬?«úâ(j&éü’ckÖaK_’õÒ7•Äü¥ëßº˛\Í…«n¢ÍWnSî»Õwı‡xG#/¯—∂/π¢\µ'sΩ9≈¨$Û%—=«Lh˘`ﬁÛ‘ì˜.a¡ôBª‘µÁ˛S’6õ«qøàô[ ÛI¸.mÕÈ˝bs⁄A±¶r/ÉJk%©¡Jã¿MŒ›ÜÄOúVÊ%S˘˜~Øß(Ωá/√Qıí	«Ë¶]8ÌµœobB∫Bwã)!Ê_˜$?|M;\&8,(ãNµlôB–Ùn6Ï}
«ìÊQé‚…ÀJôè∏†5ÃŒóñbÖ”zçŸÊ´M°ÖMH.…ôï‘[@
”ˆŒÊã≠]›sÚ¬ßmJ∫•GŸ≤#ˇØl$_uí™4ıèIúxë£Zb¢X`ÉV∂⁄¿°(<Áãú√æ≈:«.§e∆o6´º∑˝oË\ÏÌ}õµãSæ<æ+çCyX(Á7©◊π
ÿÜºkZü{f◊|œ¸†ΩzØIKX∂(®ŒTn µUÇ≤{ó|◊’ä4F:;∂ò’zd10◊ YáK«öRªóÔœ0ﬁÈˇ@S¨ˆ2¸[)ølëZb^A`èOÄ¸–ÈÂKQ<àûÊ(∑X!eÚ#*∫Œûß$-ñ!ºÑV¬Ë¢•¯ï«Ëw)yEõV⁄-Éßdm”Òi‡:)[‹Ô1∂xï‘Ö¯õ≥Ûï”Ç⁄¡Å…«Û,r∆UX[ÏÆò„.´û¥
xÅ#ñVØ† ûy‡À‹˙•ç≠Œ#Ú1∫Å◊/yâr∆<wˇ∫wôí…Zﬂ«Àî¸ö]ÁÃNáuÆu˙‹Ñ1y‹ì∞	IÎˆ–òÔ∞¶e‘´…7´ÙÿÕ‰€<PqùµN≠í‹=£ƒõJéﬂô=5H˚RˆôØ˜ÊƒîÑ’â®ƒ4s«ë<UTÔ–O;QãôdU⁄ôTcïª!mQªÓKRèÌL…n¯Z D•3¸bÊü+¨∏íJxï÷YõıËYt@›0™æ˛˚û¸éZ?ßÀÈG}©Ÿ†iúŸôÔYL–:∆¶ëL„ŒSõâ¡:÷Ë∆Ω{)π=ÌÖ»5fæ≈
æÈΩ≠»ø˙^Í|‡Ë…&ësæ|´ÚËÙS±»û»™l.k´l¶gc’†+jbuJB(3Y#8¬{óß≠+ä©C‘Kâ°◊ïb‘˚µk@±≤fe€î&'1d‹¬ö$@[c∂≠≥Ùy¯e˘â¯√Úﬁ,ıßT4.ïÚ¡ìb‰è0==ÓùûºΩC¶¶‹æî«I&f	N"Ù˙¯ ∑nHº˙wöÿRî,AÓ{»?ßr7?ÜÑícEÏ&–)«Ö#ó•dFáåÿ⁄H“Ã _¿Ü6#—fSıçU⁄r™~ÛQUœ≠Xi£é€hü.‹ª\¯ŸqNΩ!KêNù£ò¸f·™*E$T&O}·QØ ˇ™TcÄ2ª´*∂Llª›B†QE;©\· A;’y—£ ∞íVg1*5æ4˜T\Aú85 5ˆŒßÂ¢Œ@[π ëù_ZC©MT »]}∑z‚)¿éyÏæ ﬁ√úÎÜ3”/çeö°‚ësX-[Oºxë^	•∑a6Ùy.ÉQÜ~¨éöªKê§Ä%>r258MP*ê“XÏ^«4äÎ“'q†z&
>t¢óéﬂ™fË∫r‰«ü9bÉ˛ú»å◊ÜÑG≠£?≈ªˆÜﬂe_"ÃÇUU ^™âøÜºÈ˙R*≠ OãDœlãV*Gr1Û{(”Í≤o:Á"‰!óıc!5µ?"©y‘±R[Ë|ªßOy8¨¨Ó°œoT‰71…äO9„â'V‚çCuî0ΩÀöuQ2æ>u\ÊíK}Eß"ﬁH>E⁄o:Æ</%aÃòb66{0zèÒ„ks◊7∑wæ˘Á˚Ö¬ æ˛5x0”ΩÉWÍßû#·¬¡Î¸‡¿•'ÚCúÔímD,óÚóêIÖëÁ¢+  H^ﬁÄ-/™¸H‚!éw0aãìΩ≠g˜y.ÊœV‹;*{ÇÁkØ‘ë*K g˚|mâN…Ôxµ@Q&´¬Êä’®$)€´˛ßÃ¡(ı…]çt8Tk†*K|Âî!ôú@∆G&t±Ã°/◊S»)±¨ç`08*√8CôI“‰0!I
Ù¡éÑ¯ﬁO|LNQß^›ãÆˇΩÛìO0iìàQ ÿ›téËı_ˇø≈$h–ªŒ°Yœ«ŸrfÊÈl[FZ
‚$ıCÑy≤±lÍƒî&@QﬂâÉ°ú≈øBﬁ3kGŸhTX/h§›GDß√pLˆQÎ∑ÒÚ%Ybƒ2F‚Ô¬Q8§'1ñπNú`à§Aç`™©‘£@]B|ﬁbóâl¯£‹âºcVô¯%¶ÂÈ§„ö¯N4oô.±BV0O©≤ê9 §kù˛&k ò32
…µ|ZìÇÇóŒx,ÕàMª÷TÉY}¡ÍÀWºÇ`e˘Ê
R¶√ítb™íÅÊ¨ÇhÜÅﬁõè∑<Hq›∆¯öì,” E2ûGÍíh	Êîìúœ≠oº|µ˚áç˝œ6∂vgP=—à-}Å-¿8êˆ&¡Ÿ°	w–©¢HÍopÀË1A4ÛΩ ∏êÜ∞VzSÜ9X©;o©’’÷pﬂ;0ßù√˝]≤µM∂ˆ∂7_lº|±µ±µ}@v∑vww_^ˇ€Ûõ∑åK)0ñòU“ﬁg#ø)`æYÏ.ÆºEËd‹\aòö@≈OUù1Ñ¢˘.ƒâ6Ñœ¬{ËÃñÆ1êr	ÃyÂcÇ´XÈªªsp˝o˚õ/vSIÊñÍÀö!·uÀ;eBÌK/†¯cIë9Îa’ƒyaË˘G≤=Úb¨` ¸'»0ïgL.l¬l…e5R&∞gÊ‡OV™π˛Úç,ÜÀ<6Ïß%ˆÅ¸¸¿I–§ÌÁ¯Ò’∆·ˆ>P%√Fû¨4Ú™·◊Wj¸Çx` wƒåUrâÒ˙´v¡›ZÊÔìñÎ.ºzµp˘›ÔVG#S∂È™â÷%#—àƒœAJ«Ò/Á∫KIxào‚4h∆@ë'◊ÅçÃèªŸcß*zœ-®VEÔïjEÓ™+u“b˙h#2t]M+É/
‹å>uÂ¿äJ5ç`	≈öÍŒDW
YòØˆ∑,è„¬¢æ÷X*~9_/ä„≈∂}ói©|&áoÔöÊØU´âﬂm›áJQzf¢Ü◊¡‰v9ôxò· ìà¡b¢ÇÅÉ ¡ìxX(Ô¥	π2Àî˝“\§ÿ|ÇUÿÿ∫úÄ6ıcJ!~˚t|À^âÎ5∂†S ó`∂PtÇ‰˜ﬁ".ùb…îïØ“94óò«0ül5gèÙZmÚ…∫ƒQ∫_5ô(C6òD¿÷u∆!ˇc·Jˆ.»óÔÀﬁ'Urâ!ô¸ÉC’v§R¶ZìA	µN˜ÇVK"ÀK<≠˚•Ä∏µ,ºÀeö∫î6‰-Âûô2v•∆—|1–"{oj{#πrÒA3Qíuß‘P*\pD¸H Ä6"f√◊‚I0çŒ7–Î±:n70Ωüë‘=¿|:∞f,Œúzi†±2]]N#”^2<l)€ÉeW!¶»]eãÖÏ#2úÕøæfñ≈"¶”≥Ë›‡Vv6{ƒúÕDmg¨ÀKÉS ƒ·û`F¨RvÎ≥Ö+Ω–òv~oÁ˘}Ú˚ΩÁX˙‡ßÁËÊ˝ÀyóÙ_˝0o—ß≈í›ö†J“ÃC‹S·ÿÛÈq:⁄‚°±t·7“"4R¶ŒîŒ8x1f˘5Ñ≤Ê—UË{f,√7Ò‡ƒÉLâ6	°üë∏êè¸éHE’‚≠…õM‰7â·õƒ`%1îﬂ≠"©"CÆ°ó•Ü¸Æï‡ Q®Ê≤CC˝ÌI”i¸oS|Î˙∞∞UüZà»„+ë%2‡.Ï…7â¬Z¢®¯≈}/™?œFº(∞
∑+\§ñdÃcªº@^¿_·ﬂH⁄£äÈ1|
S≥π°ﬁ§<•‡ë˙~|ì<nAÚê”ÑO)z4è;)élü{GûÎ(ƒ*∆Õ$é¢ìm‡@øæ#¢Ï<cÂÆI!
µ[E ﬁlﬂÑëo¬àï0¢t£lïh¯Rπò^A&ën[	%˘Û”H%”¯}ﬁûhbv$Â4C
>≤|"Uåê7ÓSJ(˘Z|%‚âTz≤RôÚõÄR+†D4≈*“Œa«F`¯&§L)§î(ŸÌä)ÀBL˘q‚∏ëÉ%4^—`"Kœ\"aqgﬂƒë[G“à√ie=‹u±cÑ„≈ñ0˜:&Õ3oø¿9•√<
ÀëJnS‹∏]„õ`ÒM∞h$Xî§	F%Ñ(ëfÀrDzœJàO)A‹±°,+d4I‡G≤àÏt>•†ê.ƒW"&@ñ6‚õÄ`% ¸)E`à–¸[û± áê,Æ.ívYRËêØOT∏}˘`•Kˆú1åáﬁË(r\)MEõe¡ëçÂôÀY<“7!a÷BÇ∏ﬁ0úRL®óíg–]à0q'eÜM'FÛÑÉyBò∞J¸8åÒhiÃfô§≥Ñyy|z◊ˇ…‚Ü…1õ©À“)‹ö8·kﬂ¢XëuÙÖ¯~ì2f&eî°4ó6πÚÜ¯Ví8ÚªV2Gˆ¯4RGu§üL˙(EH!%ª±¨B`ÒJÅ”ÈFè‹˙ìtÜãÿäáÖ›˝DRãóA≈%¸Jò+
;ÛMà119–î˝gmnOñÒ%`ÑçcÂ3\mÀ2ç·±ä∏cxV!	ûŒÖ§6≠Iãœg˘Îqº‰	…
C„˜¯i˜MÔ≠>'=^X7
üùâs6ä„ˇ—‘∆ï¶~^7íÈ¨PﬁÆf'≠Êç‰â^y…8Sä‹
Õz–8+§U∫I]â•—¸)®ÿt)Ò‰¸Ø˝æñûYwXgÆê$àÃñ√C/“Ñ±j1Hø˙±rBÀïJ¬«˙èõayDπÒ¡ã^6&ìÄUüGQÁ8+ÀAc2‡!ŸÑ∞Poé01ªwå9-±$7∏„≈W=Vî‘ùåA`Ø WÒ'*—Iód—&ÑÓùq)UÜˆ⁄ÚØ(óQ¡$nc—äõî£òb÷&ó‘òYí=uurZ)P´ÕïZ4
<™‘0∑M˙õWΩü“qÇ†ñÚ¥Ño?L‡›úK”æb†I•c®.w™I¥ª3@ÇµJn@dó§√àúfnF`"ç∂9§Ï¡0QƒS¬	Øíºï§œV6Üat¡^xÖ∆o^lµºµCœ“Ôı/¬éÖ#ˆöı;8/Œå∫Ω∏ñ$—ÉµhyÜPùèµ9P:√"@ß‚ØX`èyTß––™Õ9õ(´-Ñå Ù‚,¨$∞œX,◊£Xlú}u19ï]6ßQ‹‡8Aª@))YÍ(¥@ÀÆ©s$ß≈oı\eZêyOdı&ªåc¨=M¬1NÜ<„5Yã7|%á∞õòÃ6›˚ïjæÒT„jTÔºÙFc˙g#˚tàelB∆ò¬ÆQ≠¨°œlf‡HãuÖíæƒòÊ}@…i8∏˛â(ÊR—ø'∏˛´#“vc±¬ ŸOŒXè8¿{å“˚díx*EN”ëä¯¿˘¿üx	…	≤ªøFﬁÄ•^gÔbÑ¬8¢¯ >!õ©Ê„ì«¨‘©÷ïH+Àºzoçœî-ƒjä–¿J.E£(Ó◊“˘¸µC&>∑[Ç ºˇ¬‚ê»¿ºv∑uHdFâ®œËÄH«9≈ÒP∞{b[ ˙vä/©j¬˚)™ƒM≠G`’>ÈR}øçY≠:∞0œffŸ\É¸åµ>¢ë≤ôZú◊î–ËßÖ`ãy+Xïdh^°-⁄É∑nëXGL=˛à´◊	Ê•Á”ÉP≠},2?•cVÉrŸﬂπı≈nø€≥HWÍÎµ"fqÉéà®‰:Õ“È rnù·ŒäêŒhiÄ3q®—q[]…¥ﬂåö[?òå√+sê-”‡4ÙO=vºõë÷◊®ŒWiàˆe' ‹:˛âC„áèˇ˚p‰x~∏ÑÊ«sC≤g|¶q}Ã¢Íhñ2}Y2·5»º˝™óñı-ﬂ|ø∏π¯xÒ·€πuïˆJUŸ≤a…í∫±i…7ﬂ?[Åˇ-ø5ïë¸x%"ÌKQôÀPŸ‘†Çç]Vê¬:ºM9%å¢""(¡ãÆ!qöÕ	%T£ìÇÆΩP∏àÎ—E%aâ…Eò™4¢/Q™X6≈ú5.+ö#@Äˆ∆£ç˛„Ì∑Üö>x¿!?ÚÒ(»œŒ›»C´l˚Ä Îs˝w7ú◊æ÷b!JÀ§B$UúóHÃaPíìx+øvH_2^XÏU8úëäs?u Jû\∆·$dà©ídr&MÒπN;3T≤&H≠<¿¬1”ﬂaÃ¡	åkÊ·t‰°˘'ähvª›µ˛∏™•ÀIÄ∂Ùå∂tGŒ∏„P’niœË≈ìÀ˜|˛ù{óÒ’˚´lQÆ÷·?SßÛäiÎ√ˆ›¬µ>MÔ_+É`f\* •ñ66¢Ñ‹,`UÄ*åTÜ.≥ntñWÊRöŸSWT÷æ€XLà}!iÔ §mr8˚2ö„b3Ñ>,ºs[˝1–ô¸†30À$´
≠“ÍÙk7b°Hô™j™%pX¥Úpeiy˚mEıƒÄM‚æ∂nØl„kFÜKõN0†¿R*VXß◊±YÓ‰R\’j∏^Ï˘‘}rÈ≈¯4
~˛ô|W<˘ÿ≠"Ó‡≠“C®"™G0Ÿ∂¥Y8èAÌw£ä∂∑1GÙ°RÚ¨¬38ZxF]Ê)[4–˚Ÿ≤-u,ZÁªfaÑó¿>-ïè1*Nˆ	${Q8†qåÆ @â4R∑íFh+ıÉ€<°Éõ^4»Õ%|B•ró∑Ä«l2µR„í‚tûRP— 0øÄo“ˇ)˝ó}f•‡æ$_ªÙ/%Ç©Oó7ù@vÆi¨évò´ŸK}ïc]ó¶⁄ÄH6·
w:Be{¶®ä<
OYMdaø,pÜü•æ AZ÷0ë‚E*º~Bgp¶8™&ZÇÙù€*2?N;±"+¯‘PQÄTƒ|Sí0ÿ◊u˛œ7µ¡§6Hë˚ãW4√Ì√¬;∑Ö€xÑ∏Ùîÿ™ëÉÍõË∂±\ËÓñﬂñ6¡Ãé}Öö∂ Õµ)¿I∂|´ÙP.aõ∞Œ†Eê~LE¬MßŸPë Ò∞ﬂt	_¨.A·î˝M©páï
VÅnÍÑ%u√ãΩb»√^A@V3¶π∫&JNÁ+£s≠QhMÃJ£ﬂªJs¬∫—ÀÓ §"sÎi\úCrGÄìQXåÇKmßÃ…ò*›ãL¸wÈf3=è*»‚V>Â8Ê˙ïHrs˝O‰g'π‰0a)π»‹t
ãdÅº`–Úój÷Y’•ìb§Ñ¶DØi™p$…FÌƒŒ£≤˘¬Pó≥6§(Ú®]√µÅQr[∑rﬁ)“’Áò&z‰nñ«∞BmÎuÉîÀ≤hÉ√Û÷0L:˝“MWk^Îƒoä’R¥£Åj˝ÆÛEEÃ∂îB∑õÚﬁ@L4p¨®˘Ñ…õ7?Ã?ÕVô„÷Â'DKΩF@≤'Q†ÕﬂQîU°ıŒ∑C'©jZÅ˘≥ﬂI˚2üœï.IñiÅóSÕ„j&ÚÔ'¢õ;·©ìÔ”◊D(%¨∂”ii¬¨ıA2«#õ:O_»=@åµS!Z¡@*¯$·~[gõÍ~*s{˜ng˚_ﬁΩõ[ˇ-@Á⁄"R<„¸ö!ìvóyƒãå.ËEíÑ$·'xÓ¯€Y¡‹!í¨ìK8Ïwãî{>FK≤Å5J£°¿!ñ®MåôexNâ:j#•'Äç4z2∑}æJv]ò@Ë_ˇ:Ù·}Ú“|wÅõ¨n$πjú≥:÷LÅõäÈZ¥ƒ´ÜzÂ	“Vz%&~öO<Vì0ê≠'I¯T¸™îU™	~UC√T)3V¡6ëît∫X.?◊jc˚Õ‘±i0ÿ]PÀ2.¨Tù~ˆıÿu0™oSbøæS ’_jAçR∂†≈–he˚&µlñ2¨Qò¶?,≈ÒÛ@M¶Ωù—"4TŸÊƒπâ∆∂2÷€T›ÚH‘«?˝dz[ií÷ËsR⁄*Ç§ßR⁄Æ|J[¿S£“Vïb1Õ¡úe¯‰	7ﬂ<Óùûºù±öW≠πö.IdU…ªîºg
ﬁ,¸ΩVΩ´ 9°Ùﬂ∫-EØƒôÜØc÷gß·ïRbH˛T!⁄øØˇ∫RNàè´“U•Dh™“ÕÛô€´tÂåÕ¸Ê√;°Á%Yr„ã2°d°úQ«ArÁ?gt8nt˝„èÙSd·_1D=Å:Oã•i¶Ãnß@…ªy§ÀRÄπ
ò‚!	Ä_ )R¥Èt(´Dùr*¡‚®π9^ˇ;`ÿ ‰¬¿≈ÇÃ—á¡√yúcﬁêrkä; ö∆iFñº¿qXÜA◊Is¿êÿ WEG °;aﬁ¿µ≥m®$k}0ç`ê™	çj.â¸ìlÕ¨˝ƒÌæyä]g5dê)*§âBÙi‘M√âk†„JNÎí≤≤¥R®÷E£^õiçí.
˝\[Çª|¡t%ﬂ%]ñxä∫…|ö˛@5@Õ–ß’ù⁄Ó
^õ:bö}∞KÒπ,‘>˝”ƒã=	=„Öä†m =”≠ïûÆÏ2#ê¶NDe=aër∆®(»mÈË}≈	ãólÁ‹π˛˚)tDòqjò&F±BBÌß(+ÂM]é©:Õr z%n®TÉA:œ·(ãA“Ô©-Vniê9˝ ™ÀÌï◊L˘ÎzR.;êŸ¥/ñ&ÙÀ≤:
-ï›»RcPØûÖ3√5ôπï|uﬁqÖ®Ê…JÓ®R≤X :ΩÏç`ß¿UèíŒbIàí¥™ÈÜ©4¶JñMü‡æîó ú’ûë©ç1ú8œÓïª¡x±æ–ŒJÏX'Ò„í[%¡^b6Ïc§|/æl]ö«B§µ8¡K$ü∆∞¡·•R<ıgäQ,1âﬂ-∆æ0;U<…Y®åFª£N.mÃÓs+ÑÀ	‹w0Ú¡áœÇzÂ£˝F¬fA¬Úıú´•œÖòeTç¥…ı/¨ƒ`ñMÕsñ¯∂$ƒë¬ˇ’ôH)nÀ…{4ﬁÉÊÑ¯wÅjyyá$¯f˜Ω)≈˙ˇ   ˇˇÏ}€nIzÊΩü"T=0ã≤ƒ£ÑöERcí®!’√!%´RdŒTU“ôEëj.Å|·K?Ç-ﬂ,∞¿\˘¬∆ﬁÚMˆv¡Ò«9"„èà¨*v´g&g∫ªXïáˇ¸ˇΩR,{§°V≥P+{.g†T	€ÁóB•s˙ÜDA=æÁˆ©Æ$^K‰(ª˚#–@Y6Ù”≈üólΩV‰®!ìAÇ0√FÊ]"uIÈA1—¢ô§YÃjFºªˇ≈rïA+cv7ˆÜ”0Ôó^[O«`Í4√,FîÍ˘(œ
71∆Ûﬁrƒ¢:}NuÍ30y“i≈’Á∆(çGÃXÒ2{≈8œﬁÔΩﬂ\Ì›É∑/_æŸ9‰k}ûùC∫*$pFÑÕhÍa"Gß≈¥!®”	°3ë¿MÃ ¯§†≈êêÖ˛nL´QÒr›G•<loÔp#—Ê@©åE3EÕç√˜-˜ÏŒ™pTçvYŸ®fHá∆}u¸]kM$Óü1¥&yByËç=ì±®õ=QÖAﬁa%]ˇQÏQEa‘]ÏM ÄhÌR¢EóxvÙô}Ωı£äXL˜¯˛ﬂÔæ˛·’a\ölÏìñªgãÃ9Ù«@h«D”?°≈¡Ëpî.?7†Ûgëî}Á‡˛Éô∏≈ÕXd¬˝«4Ö∞˝ãy|ÉÅMî¡ñÂà
mü¡m]}!<∞I≥πg;|ﬁUÄﬁœuõ,*
ûV{Ç¢í√¢ÊôÀåö"4™up!˘u1i”Å?úÍ¯ÈÊÁ´2<€“˝ƒ£´Äƒ6¬´6H=⁄r¢¨Ã *;∂ fÜI±PÌ3g€E"Å∫êÚ=hΩ¡w∑h∏ùH8BÖ’Ç◊Õ}ì]0À˜Q6ÃH˜]YM21n%e÷„Ô?yº˙h˜Ñ0=€‘*©¨O;§Ü‰`ÎRI!
Mj@%erbX‚”C•¶·nû¯4Ysvx3ƒRÑ8ã”»∆óÆﬂ'D»@û8®5¬≥ñÎ~Uáßô/ŒDë⁄ê„ÔﬂáÖÜƒÛ5ç'˙·39ïX Âä©€ÁÏﬂ÷ä †ùÌ@5å’1w<π6G-	¥$7áoéXŒTÑ&MøX>¥ÇóxNÆ"≈À3â<Wÿ:lx¡j∆*ng8EëG–BN©÷dA’¬Q˛C¿œ≤#{YÈ5Lxó/£_ùYCŸÍ0ç 
u÷®<`W
PuüPK3ø¥ΩYßéXˆfY)œUÑeè«ﬂmll¨¨Ôû`π_÷Êû—⁄™ë⁄¿„Àz¶é|g˚uŸtÜP)zB€7\ì…lÁ$aE∂â	K<ˇï‘◊Zè»Ã>&îåK™{–Õ¶‰ít˜≈‰ÓÎÁ|ÿJLô!ÛœJ⁄ì4'Í•»≠å=E´¶O÷{¶âë∂Cˇ˜ÙD”®∆È5)YìÄq”£KF≤V^NÜ≈8gPuâIël¡®å∂{YeP∏∞^"˚£º:„—ày˘˜üï<_1#eÌ®ÿy¯π‚=EB÷{‰e1úX08X%cˇ—Cf6Ú&1ƒ#∂úToëÓ1›t:i¸9ppkè)ãº:,ú¬§VpJ|À8)ŒŒ! ª5€{ìU ”˚ºˇS1),ı¯ªΩG{Î{´'ùÌˇ)7øò}“´|pÔsæó◊£YßΩm]7‚õﬁ®ú_è∂‹ Ï$yD@@˛Áﬁ ∞¯/JZYKå@‹Öoh{–¢®)tjVEß°¨j'»£˝˙áD~qAÅ3úñ◊H¨ÉxyÒP{œDØ˜áóî∆´	ëfCH°‡ê@…›‚‡@Ú¬6“âH±Ë√H˚ÃÃØ{ÈuTC ˘êûÂÈGƒﬁ?™%5àìÿ ø4IGh€†lW©ìzCÍÎ∞√,‘8Æ@Èm4≈¡!ßJY£«Ër”Ÿ÷':¬pxç9ÍFx>Tì> /îÓï‰d/Ω'ÕèÌ§@#V˝Ã\∫'Õ˙Èj¡yTZF$ÚáŸ∂	’”SUµQOxOTÒp!,[O∑…ym™≠C	û{^Ôˆ^6™	|¡Ω≈Ë¢ kHÔ \ÆIû/3®Ò¯ü∆g€UÖ˝ ø∆	29(…‰ÓèìÀ!Kª≠À1î|áûe£W√Î$4·¥π	8çF‰ÓﬂjˇÍ®}D€Üµª∆Aà¬Jær@*≠3ä^êa	≈Ë©JY20ë”M8πîÅ?•.‹0"Éo˚~Î’&Cb?ív“îîÓ‡≠õÒSô˘u®¿◊Ãq=‹qÖ‡jy®∫1’=^Öoj;$ó†ì˝k®ƒlM8• ñ±DÏ∏D¸ga‡±5ıB⁄Ö£_ÇgòôkWRv Ò⁄âπW¬Â¬¶ﬂ ÍÎÿqÑ¯Z}R∆\˘¶[FêÖ˙ÆmëÀ≤ûË“íŒKÙ⁄sﬂoØÈ¶lÎ˝≤øÚÉE9∑*Ïa#4@¸¬#x˛Ê^>…ä!€l=¯¬êüE∆Á_†∞QTïuüﬂw¯;æ ü(€C$€◊‰¬6êÿ!Ω¢>‰∞°œ…¸><ÁÄ{˘Ár»√†Hì˘€ëé©¶øﬁ˙mrI>{%Pv∂øãét–´iøywy”âÚ∫%À‰fú_ë=∫C∫ëfÙxûyprx‚hJQwÅ
/ ÇWpæláÉ“S˜Üùﬂ-¬Xá˘ÙZ‘ÃÎ5~[˝…}ï9õ
“ƒ=ò\ëÒÒ_ZÉœàhñNCπig˚hˇ˝¡!Ÿ€ßˇ?zˇÍÌ"«Œ’'óµàr|˚˛pˇ7?À©«¿úÎù¡®˚!<ŸÀ˚≠ºOØˇÑ˚ãÂüTx¯x$P)eE;Pmn/ˇDŸDE∆ô»CA¥Ãê‡Ù“’ütö˝ù8—Q«bùËlÆ√\Æ≈‚¸ßc›à:˜À}÷⁄=›\ASÙ©Tn}(3ü¢†DÎÀ}ªíóåçL6ª1™tœj>≈Òâÿ∞èˆwÈë‰¡+tA€µ	˘KÔ’µ¢ÈÁmˆÔ5bàuê~_f’ËÓkø»:€Íc∏aßÅÉK•ÌlÛˇÜ≈@ÅÒ_EÙ¿ê®=≈ÄH“VﬁK}á9©Ä%–'◊
ÜrÅ≠ËÒˇÕ´˝˝-]>36ÌW«˘ö“Ï”Ív˙YQjΩ::úæÌL.!®≈ÒËpÄõèHV•Ik@ìVmπÄ7n)ç≥¥Œw˚o˜(I¶‰"ñ$@öJª·R%ò
±≥u¥ˇÊ√—˛ªù√ù›ùŸŸ≈eu14í|ƒﬂÛËnÁ›·¡ÔvˆTW ¨lŸxg+≈Õx7f›XÎ$Õ£≥√˝›éåwÚ£:Œ‹Õﬁ˛Ô^ˇ@óÈC“Ê IA–õ„PƒzÓD'˙Ò\ªSæòr¥C∏ˇ¨ æË˚Ÿ_è|I.∑ΩÌâÇïˇÄÌø!¸ãªæ˚ß–B}<oˆç¡øfù˝ÉoFÊ∞Ùfg˙á˘uf-∏Ÿô˛a˙Œºï9<“7û]c}ólõà∫zºˆ…aû’•Wëqy1‘ƒ&m15-3ï+^Â«oJksÃ0~∆Ïµ≤∏Rò¬ çÀ`ÊD&J:.¸}¨ãÚ¥Œ´œ<≈˝øTS=”‰v¿ÁB¯–öf7˚˜ír´√¢≈q‡O•≥∫”bÒ¿wy˚#5˘†’&π˘¶%∫Kt‚Ò⁄,¨ÚôyÕ©DœÇ¨U%Ë˚8Ì”?ÀññÙÛÅÔiìÛz˘n*î/ù3HW∂m
uà˘»á¬Ìã\ê ≤»5Mâ¿ˆ∆ì\§Ë;-õÈ≈eWDﬂ(E¿ÁvÙÀd”/ÔÇÊKáyˇ≤Œº$äYπ
UyÅvZ%’·ŒIÛãí.óÆ‹˝ÚÎW	úŒŒÎ©G[ÏsU^ÒX·¿ö)$!TüéQàóeE◊6/.†H]6ä˜d•≠sòå≈´X!ª™êÒ°É¡Èzﬁ+Xb>3∆·`D>ˆQ6ÈüCﬁ<8@Ë»Q8f¥b#5_∂f¡˚a∆> B{®Ë1Ë†˛Ñ~∏+Õ]IST+„)}¿jYîøæJûe:§¿àLHjÙ∂@X#_^Ò,G±SQù±˜©*G]–ÂìÆµVlâPˆ)¸a|	R…¢\µe9Ã≥Òb(¨ìwÕ?ö–µG!#)∂…JhEû;O˝æ,∆›Ö%≤∞zjãW∆î•Ÿõ‡¨[Çø˝ìÆ˝Ÿ,ãœ≠óg%óóóË;£}ãrè°8V≥ã-ßÀ•¿sˇxILÊX/zÚÛáÏÇ>˛9go≤z⁄|ì-µ*Ë∑ÿ∫ﬁ"+Ó˝Rñ(5òπÓQQ6#ã§∑AÙ ±E¯¸ÛCªE¶ÙÄ‡SÃﬁsÀ•◊¯˝äPl≈z.x+‡8è6 ﬁhH¬ãBì-?±l$êÓÃ¥kŒ7> t «B§˜ËœıtòC0l1**≤[é‡xÉbÈÕ∞®Ωõ6–ÆÙÄ“çqI•∞wUIi|˛<ù|∞ïÀÜ5]>ÔÂ˜"ÊÎÊ#%¶õsfD¿d¶¶5õ∏Ù 9F1%¨ˇêOdÔr <+?Iò ÜUûæï£ˇ{•‚ï¸¯@å'W0ﬂczNúF3Yv{ãµ(xòñ0
sR·x≈AÀ3≤fµÊøï√ÄÒn[ŸÇRX£·5uôs5wV≥∫}î]ˇ6˘fã ¶˜SÑ±gˇxÂK&ÆRÚqÔ.[#˘~¿∏qæ±”≠pŸPæØ˜eVO∫ù˜.†1‘¯‰¸,#øø˚
ôŸàL ÚrzÀÁr¯π ò¨ŒÈ„Oe'∫u–◊«^çÖ5Ä8«ˆéÊ´¢Zœqnw·-®Âì¸sN	`πêˆ¨°$wÇœ$∂EZ–=€èO™À|Iä/1Ÿ¢ÈÔ>AõÄÁ*hücL8–uº“D—«2kóe0Ú0+–Ç?Á[÷◊æà$¢D√…EÑ	¬ç◊t∫nÉ◊¯É	3\ä-∆^õ6à+¡—åZú2“◊oomB˜Œ÷!ä
Õ◊˙ò$ﬂ⁄H∫,„h∏àÿ—lø6ëy⁄wãG2ÜÛÕ	èk~¨ælníƒ	K·lô˛ØWÿIJ˜á‹aëî¢X¨®∞bÆïñÈÁÈAïœ‡U2*€yZÒrã¨ﬂÄŸ@;ŸqûK	˚3QD¶=ù˝Îã!ÿàHIFÍTÏ Ï=	˙ñ$ôY^ílPpLÆ¬ÃìÑ«°˝∞,©†,˛xNô	l¡∞∏¨âN6dB¨Æ¿^MáÄv∂J∑c˚˛pØ3
\‚Uˇ|Tü•˙j◊\~[9Èå|‡ôÍöÖHÄ«ΩﬁŸﬁ·€*S\wHˆk:}*•‚xÎŒÀ≤\O*%$e+o¶∑ÂY’?∑–gOkêÔretO ãÂ’ákƒíU‡˝ÍW2 qÑ·Ã%Ï[Ë]^Cë:ï#9ï,c8ìSå@¡Â9Ø√Â’5 5‘0Óæ Ì∆R Ó÷ëå`¶ªW¿ƒ¬µm¢º3ÄÆ¡ócf(i∂ûN£Ú◊∫@í˜)˝héRô{ø'k	 á≥ïóëõ¶Û—ñ˝7]'ù¡»P”yiÏ±¿CæCÚ#DË@4ºøIÒµÚ ¥_Â¨ùÉ$∑Ñ∂µMFom˙¬äõn#DZíÜ;L˝¯‚EΩ"#<§ØÉ!Ç‹p„á¸S¢CËo¿¢¬?ﬂnìÔ…jK éàÅêmZËQ6ºü—1äq[√Á]éfL:âaõ è¡ÀÈóa„ÌãûÇ–”
ﬂ_‡≥QÄ¡%0!lë€ê?âØü?fÙ;8˝}ﬁß∂Ó:ÔÅπl·íãpw¬‹éAﬁRb.RÌÿÏ≠5Zå{SÛÑ/.{≠Àj“ÌfK‰îÌÕåu–≤<?0oSY¥{ {∑√≤ÈâﬂMÏ	9Y…∆'eMÙÂ0∫-TE+õÅÖ›\Àˆ€||9í=ÂcÄ∞ %´«ie ü4ÍIToñbµvX\Bù
Óñ$£<ÑmØ1°/Ø™FwÀ>TJÄ-Ûæ¢aÇ∫∂DÙ ÂÁÄ«/≥|Ö‰ztﬂBﬂ·Cµ^Ï4˚ÕêÌGö4.lÆJ öe»¥Q:ÜâÑf‡ Ö∞¨‡
Å»®{1™ÜUCß,ª	6/Dj‘\b§ãŒ\¢HáA CqÀª!’;t}Éñ·h∫ß0¯‰ë·¶ı.ñcÜ`®6¸%Ö∆DÅm–!lÉ¥ ¨|ˆéjLŸYn)≠‚syΩ∆Lô6ô±ã)Cqî57-óM¿#¢T”®åIxé,Á{◊	P#q—ÙA¥Ji√0ª®eorûgD€öTŒ40bOÀ	¿'§±©Ó,6®°≥˙í*Ë VZ~D*≠˜Ÿ√…˘‹{1yr;,ÖﬂN=±kä∑ë≤Ø‘Æ… 8r+œ∫aLô”´AIb@r⁄óuËP1∂÷j¨3¨¯Îíπ‡D6¯70${Œÿ˙æ.N|,2∫ü~prÓˆh˚≈∏è˛T˘ïqî4<õúñÉ/Ê0)≠+î÷Ò!F	nÊ$¡¢y®p,c"˘√ád7ˆ/áY≈‡Õ†Bd.pÃ2Háe^·ú±“≠ÀQpq™.{S&BÒ
(K*∞Æäƒ5«E6#çﬂ¿ïnÖ8≠C¸∆¢p≤áﬁÄ˘øè†&3Ω˝:t¬È_y∑[_é®z ç@TA˘µ°k/E;`°"Ê¿ËÄxJ†¢)·ïd4B§ÕÛ¨~s9úC#’ºä\≈⁄Bæ
Æßdär1∞d¿úŸ ¬!4>H¢nGâ÷qY˛¶±on$ﬁœt£&8ª«x¢ ™»Ÿè4Ê¡)ëŸË+º‘û7˝˜… ä«Ô∏Û•à6
aOÍÎˇˇÎø¸Áˇ˚Ø!7ﬁwK^3J3VU¿„o©4!Ôä.TPß•ë¸ƒÔÿ∆Ct:”ñ°Ÿí±Òn˙≤ÙÃ´Ö»e≈È ≥-Á8òe&¶ïŒ¬"<øíìOs¶TRây∏¶8Q" éz˘’ç¡!û}OV˘J&ú öü‰Ãµn=lWaN:(V£‰∂C`≈å˚Dyå®˘ÊÕDLrLıô—‰F‚›≤ö€ÇHêW„˙ÚU`ÒRﬂ3b√J3gDC∞èv¬3èÎ0”dcIw:?ü÷Î¸<‰v∆2$ÕÀ√IS#HπópıYè‚å1∏¥Á/´øå˚D¯ˇbòÍÚ‚"⁄8øa¶T4≥ùÑat}Å√—V∫Æ"°√û+[ê§üüì“Îı™bâ»Yÿ2áuK  6(ØØIı%y…Æ≤bBO=ÑÌï˝Ó ˛9]"J◊Å ∑Ó‚Æú…#º/ï≤ªyU≈Q˘ı+VÛ}¨¨∫ù}˙íï\E*~ ~¡X∑:K⁄OT“}AX|c§~òÖOè8T9Q^∑Â√ßõñkDJaHvòW^‚@]Â∞É_)|∞ÅÓ’Ÿﬁπúî£ªØì¢_íÓÀ˝ódôº´äQ^–ÂñÒ¿d√òaÓÂ
≠pÄNcû ﬂ0ôt ¥Aoÿi£Äƒ‰-Büj‰∆Qm~]∏%]˙ªd›∑‰rLß‡wÙ≠Ÿs˘ıEQ}˘ å=ˆäÆ‚§∑prˇYo“j‚0T7∑Mi≠—•>~çhÿ*p±F8æ=U©å√Vñ˜Õ«õÎ˚1˝Q^bŸ¬cª≈d$üÎ¶≥›k⁄l]„T)]hƒ"Í8¢åÏuaò{ó*Òí8ê»iùÔ∫ﬂXVúÁ‰#ﬂø≤æ∑∑¿G¶WﬁY∞(€õªˇ√n(kÆ9/Ç>—ı§¢8õ_!≤a<›6·ß≠Q§°_ÜHù®&˛"ƒÍ`¥~≤>ã8¥/Õ¿7Y∞êìæË4~ﬂIl!æa˜h+«O%¡”ÄU5´Í¸’x‚˙EÌk~í|c*∂ÿ8€…Ú≠ÖÕõèWÀk+&&ØrëwÖUÃ≈ËjlZÒü¡˙æÊÖò¡∂MKF≤Ü#≠πÇvµOÒàÔ»q$S.˜a 5zí∞x˛°y·ÂÅÕ+ñ¡Øª…‘|Éq3cdêSˇpßüî9yîªñ≤ÀîrÀ¬o©ÊKﬂ®b"	ŸÅdêàÒdÜi≈|≤Ï1$Î©í∆\µ>¯%?2-íc¿∆6HÛ‘EŸ´ÍxÜwì¬)+b dá£y´æµÆæâÃ1ShKÀ´<îÚ0á‡ﬂÙ ¥Í´:—ñÂC¯ÄÌLy¨ï◊T≥üHe*~·'éß/b¿4:˙nÍ é∑˘’ÀayîO£´ÉÖ∫räø3L⁄YõÄ◊˘«[,µ∫y <Ç¸ﬂ=ˇÓâ√≥å+P≠·Rz–â∫U[wã™?t‚˜≤äº»äÎ,≈3ã3Ó∫üw£Mπ]BÂ–¸{)u6›∫Ô0xù .û4˝Úb±[HH/¬$õàÿ1ˆ8?»Ê2˜OÀÈ4≈lÃŒÃ
xWï˝ºÆãÒônΩ%õ‹∞ÿ§ázßØ&’'T“9±Ræ8 VëﬂßZCAóZ-‚OLöÓccË(æà∂dÜ}!9$‘yôâC‚	ôúàâœ’D*Êr}AU´…˘Ú&U6ïÅAjÿLu∂#z}L£Giì¯c˛¡ß$Îàü.÷G†5tÜèe[∂r0äJæ≤äâ	«dùáÚøW ﬂÌxëÆº2W!¶ˇ……Ëûl∂/vøy…Ái4G©v)Õû£π∂¬‰Ú›r¸©®FÙ Ól∞úÀÈUPJüäJ
í∆É÷Hﬂ7–rfBSL6˙`UÌgÖØûáL ºäT{GzïÁà:“(Ëî¢ë4mS≥co2Lƒö≠ö
Ëß∞Ór≈ΩÁ≥¿{,ï¢ÖIk´<O◊|W˘}ï’ÁÓ2_˜áó≈¥ÎÏ'VÊÔæÃSÎ)M‹ò¢È‡™™◊-jZÛPÛß++yÎZS¨UUÎ@„0 5√|›îb◊&†à.|˝(PÌ⁄ƒëïØS™]àDNòü(…˙d≈ƒi‘«ÜfÔ˝ÑHÃU{ÕSõ»ﬂGx}lwátÉìú&∏í B+¸5ï1‡H¨†≤n•*ŒŒ'ÙMîËπ0<∂l:‘≤p5áû§{ÆVâ˘åÒÂì«¡ﬂ8HèÎàY¶ıoÛVπ1°II6Î‘h®CßˆnÒ’SÚ$4Î∏«’NöY¥Êû?'§Cπ∂Ê-Xçî·Ù4Ë)•Ó”é%˘–™ÒˆKg˚-U…Lçí+ÓezItµM§#ì8OïìŒˆQÚ:Ì§§6mE√∑‘LC‚≤Äı=íVôó8≠¸â◊ﬂG∑€[JÌ˙9§OÀÀIt€ªJ≤b´U¢)πmbÿ±âzlLîY∆pÚùM%›jíPj|’M«®œ´b¸ ;FHKŸä Òêæ}ÄHG—’e%›vÀQI>]éF…8H1ﬂîœì+Ç±ÈÖ	êπ'@ÒØëÍ‹2Y"£A:Ï8¯‰⁄(ôÜ∞Fí}ŒÜªÈ"“?á£íjO≈'¿≠!;wˇ≥r∞Áñ†∆	nPöF{”é≤˝∂Çc0Œx≥$HŸR?„ËÃÌêÂ≠CÍ:◊ÃzÈeñMˇ≤¨F‰eëıñDB˝kQ∞c'·¨¢
;¸FÈçÉ-˝Á/ˆ„›ª≠ëWùåFáÊ Üø R≠∑2!¡"cÄ:€ÉÖ•‡ Öúb˘4UW‹F∫N≈m d'¿”Ö{Ãt∆'.]”bÇÊRz—$e∑î¥;
J˝ÚtÑ˝‚\^ï¡N?∞Ä≈)?4ˇñ∑õÅ¿Ób@⁄<J©
–rQ›˝«5ïòi#ÓW-öS¯ô*~zCGÂ˘≤Eì,OfÏP≈Œ∂˜ÎÕ¶÷u«r¸∏äﬂÖ±iÈC ˘<o"&‡¥—5 ASÀ»ƒ~˝ñƒù˝Î-çÊZóß¿36yŒëA¸ÆD–Hˇ§àïêÔb‰*E¸¬X˙Œ`@ñ–;+T»)ŸÁ§¿P/>$´=Ú2á B&~í˛eU—%~°íã®EjΩuıQFÕo`°π≤jKÃ≤ÒŒ.8H„Ün¡¿Ÿ´0ˇÄ•∞brŒWÃ0¸2´Fw_˚E∂¿13≠6å_ΩΩ⁄ õŒœM@E:wkrÓTÒ(YÚ¡ãï¬ÁL›{òˇ£(sOh'	ÛëX•ÖÏw•cØRç*=π;†Æ3ÄÆe];kH<Zù√˛6Db õÏ¢	çjUïëﬂ⁄Ec‰∑ûêw™j¸$ÀˆËüt*)oá ´ŒüØî@Q≈jP&ü´ﬁyVw-ø’¢?5πvP-L‰5ﬁÚ°»V∑q¨N¯™Áÿ©Œ }<ìãzfß(!4e!œf	îÛÛÏü¿›zK9uÉ–Ú@í≠ﬂ€ò\œˆ¯µ€¯;LjZÌ0⁄tÀ≠◊|!˜õ¶Û"¿,ˆ$≠ë√Î’o√a∞8íks˜Q‚¢πöT]/Ü¨≤E.º«Ó¢˘ÊçóQµùXﬂ=’®ø¨ÂÎîPEü∞åN∞‚≥WKa†‚	NŸÂå∏\’‚¶EΩœüÈ≤î†Ê`6zPÔî
'ÑK◊`p™xÂïy+	Ãº£ü†ˇ!ˇ“§‰qÍn,F∫óËΩ¨ {€Á[T¢cES5Ew∆€úAƒÎW(`öıv«Ÿj,ˇÍÜ=y›1«^€søèåX≥"Ó›«gÍ£lñ·Uëcô¬¡ølfè?|ªp{‚Ìﬂòk˛∞B1Û‹l¨B`®∆ äVMJ‡|N∑Ôÿ´U‡´∆8ô≥tîˇ“Ö`åÜ.†ü≠•>çÚ7æ∞Á¸‹ƒ‹Ó3·'¯ñ∞ü¡9¢±‡¸ıÖˇ~ΩÊcs.˝ÉÀ√⁄|4ﬂKJ‘(Òq)ºµ(ÕVPÑ:ìM§Ñíp∫˙∑≈Ÿ˘< ‰E6{‹M‚¡&Iv Ò"ëUMØÅù}ÿHBtΩÚÖ|Å&˛>¢€Å≤4Á⁄ö·˛¢0ü¥ΩÑÍ◊ò°$å∏3Ã´…˚™»∆g:†–É8-"à” EÑ1Œ©£ÚÉ„F¬a"Qî~{q3ãOÂñöëÅ`˝l_≠wãßﬂ®∑°MÃ8ûì˙ôœÉ9„éÒM⁄ú£P™fp=	ˇ™¶,≈è√Ø=™Êˆ'ŸàÆˆHmÄ~9É7¶≤y»,1E¢£¿ªs	6Ø≤¢≤ØyE?-'Ã7Sía6`^Qp‹p U≥öŒ›◊Âa”Ö√Ó0Xd»pm°”Ö¬∂cœo~‰óoÂAé¶_„¨V_úÜı/m˚JÒ§D|‡‘ıÏ’Â(Á:ß∑b;Lt«´÷Röy⁄sÍ≥Zw0Q·äÇñ™…b¿=∫’[o|	¯ZYp»∫q
±&äL€YüÙAˆπ*Ø‡s∞æwÎ™8ï8Ã[JMÇàwı∫v‹ˇùÌs1RÍxè©0ò;Vr”‘yÅ˘ùõscq  snºÉ_$â¸ﬂˇ˘Ôú03sº≈]∂o¿∂G˚ì'≈ÑÁYÏM ◊¨‹!`äË˚æ8\X¥:ˇÌd@v Z,ıÌÏRpP©x:)U&RÚ˝SÍ°¿ï^Ê.´–àMQ‚`âEÀÕK°ÛqRóà	b
ÏövÑjM7üÂ¢ªy–“û◊±›-êˆº,ë=›”çöŸ˙ıûñr“]4+B˜z=óù,…e;I¬9Òå¢^ˆm¡ty}tB£T…"!AmÅ_Õ@.HW*˘aö∂¯1·=–^L§Ú‚qî≈ÀAWS≈ı‹†
NJBÉi I€ôcæËòGÉÃn–?˚óU]VÀ„rc/Ø®fà∏±_∂4∆åï…¶$Y#wD’0å£Ωƒë^PÇõSHgÁÔÓæí¯ƒÑ~U˘∫+Nèø–±y≈´˚ª"l'∞ª˝ ëV—ülØ£I6d’@{R˜Ñï1xî|”+M5zÿÎ [nä„ÿaÃ≠ﬂ©-‘Í∞⁄·áF}@d≥ Am6ïs|⁄9˜B;]Z	ßY$(≤xÕ≤C˛)«≈qfØΩMfXDQH¶HO∆;Œ›Ò·§8j<üVG.O:^µ]‘CRM÷î»áÙP,~å).*ÏÚ≤⁄DP®ZÌ"¶[õlyy9pz„Z"¥»Ú¯<hßÍ^ÿ_Q˛}jX|ù[¿qávyÖâ´	T{—£ˇQHµ¸Ø(vî=j*îÏıE9æ˚„Á|h•⁄ö¡úMAö≈≥Àcò≤1Ω$Mi°J0ØCì,
#DH"≥J¬PÌ√›º0åp*“E≤b;åÆB}¿D1ÏPüH±Yµ∑EWv5û∞T⁄V3µ/∑ª∏fÎ_æ·ˆÍãa1È.ﬂ⁄¬‚Ò I¨U≠ï≠∆nï
òÏ:‚Ñ—WC˜í-ò¡§ÿHTø"™Ÿ‹ï2?€[kbÅjø)`ûSöƒXÕ‘:SBß∫—&è5úR7rR¨û4” Â‡∑DjUKÙwı$YΩbo¿%…qyUeS0cßÙ'Ë§™P1Ÿ•‡a¸D(G~~ÌÏ"^!‘˙
Ñ˘◊•;Â'|°w39Â*P,øó≈Á·yÛV	Mr••IÚ~∑Ä[I‘„c≥≈x^#tú±	ÕxÂ=Ω√/∏ﬂ4D®∏5’µf‰·iÅ5›∑¶∂∞z~»–oÒó≤\j⁄ïÊU§1√nCB¬#›[ñb’jÇSä5Ïô.°‘∆ú[JO»ƒ∫qò’Ñ
Ú◊A·m‚N≥Ú ∆ø‚uÙî;™YñŒ-ü.≥9ö6\˛√äîµ-–áDV1Ñ)§	ï¯¢
‚û0øÛƒùvΩ(æÄI’ÁBıX«9‘fwuﬁÕï;«⁄,⁄◊™ÓﬁõªØ◊=¬«œ ∏.˛b~
‚‹s_ÎeÔ:ÓÏm¿í1ë<më¸	‘» öñÁ^&:Ω#•ß@pOâLœÈØ©⁄£O?IÉÄçSíy“ÙªOU∞ ß1àníº-I®ÙN˝KQ¨»L˝âZ¿ÜgÆ∂!=Ìa´^{X‹◊ô‚ˆû±@—ú yÀÅ«…-<Ù íKXUöºc®¸0„›œ‡ﬁ6\`ßÈŒØEÜ@ûX´≠M£ƒ(Iåfø∑¢ô,s) !åxµ1ñN9  ±nNı>‚jâsÇ÷˙»$Ú´|<"à8ó¯U~^^Å»r√†mbËmG∆„MNÃ}1´xpèˆ¿<äı_±Sá7ÃÅEXB%Ω}˜s@∫´Kv5≤LV€{’[Ûâ.aÑîX|FU∞6æŒ’ÚrNˇâ⁄X÷¢¥-≠9yÅul~”∞˙¨7Ò±Èˆl ¬ tQ¨∆œÃ>·Âπ¯uCY ˆ’≤à+ ¥?8t~ﬂFTˆóJä?j/«ô’Ö¯i*∆]c∞K÷’ZM+¯sÔ4ÅéËõ8¯÷—]]cg7Rï»ÒË	eáéŒ9}¿U#›GÀÍ¸πÒwW€ƒ¯◊ﬂ:∑ÿ∂ûT
ÚgÕ7~=æëdçH:EmŒPÎ‰ûüÜIRòªË¶gmLOÛd|ÀŒ≤·£€›≤*Æ9jàÅÂß¬”ÏÄ5ªZpj·~—C5§›Ê#h@ñÜ‡<?•©†œ»ãoƒX~OËÁñı™º-a∫e‡ÕfoN,£ﬂº¶ÃµX,ôÙçÆ9¶—‘ëf(7bÍ◊}àÖ¢≈P.ÇÇfÎløÕ«Áó#ûKSÀ8±AŸ»öÈ°Ìœ.∂ulmJÖÂéK÷"‰YS29 ¨\ùå~IönX7ÿ¢!ŒœW∑$ØJ"0C„k·∂uÆJ+5;étâ·œï‘":≤l$4ˇ∏JL·ﬁfê‘Ê°⁄Agiêt5Ww7˜Ûa£ú
ùRYùE@4„u2bï5@9W÷cUØ«Zs=ÊQá#}-åBU‚ìÍü¶¡‰∑å¢àoÄv•;Óød # òÎ…/] \í√W7√C}'§ÒRÌ–≥)…|_fTîVÄ≠s ≠ûÌ˘~WÂuNÆnÒfœy†∂IkÄk˙o*•A[√Ó≠˘u1Ij”LŒ‡‡ß%ù“Â˛√¸”dyı··z+#ã◊Ïã¡7πr˛∆GºfÖïFÆq±¿d0µ,_-S>ŒFJüG≥9æT}Ÿá√± ,¸’1àõ·Í6ìWUY-òïyc-¿ç™t¨yßuüù/ao∆Ù]ËAg‹Ü€/-¯“ÿ}mÑÂ§Qg[t2¢c£öõ•È?BÍ@>{ÿ8∆q˙°Œ·∏C»É:ˇÏP1üv¥‡PA;º—äˇ®≈pÌÁx∏·≈¥?~D∑u™˝#G–Aœ∑u¬5l˝íq2õäïq–%nΩÒ@√^˜‘|<Ü=eêÚHØØ¡!}¿8Ÿfï{´ÒX‡⁄’ÚÍ# òËø‚5¸$ƒY©?PÖ¡]÷ø•ƒy5	?‡tÂú„p ¬?‘ów_´p£õ ÊætÊÔ?yº˙h˜⁄y‚¯{™©ÙÈ·ŒÃK_°g‚˜Tó˝ïTŸ†gïePR9>€FN	/8#„Q¯≠œõ‚„>h)}∫‚äÙ›W⁄ï¨>,ì¶¶:u> \dh.Ù˝±ZB„ÙÑP©ë„Ô^n“ˇmú8∞#≈Féø€ºøπèEå>Yë˙CVÑîâ Í/„>	Yõ Ñﬂ(xTˇ§˙∞_eWY¡∞%i{{eø;ÄNó»4Z/,ºK/ £æ‡A&ëu;ÚêÒSp˜G∫'mÉÚT8
:K§#¯kmñÙYΩÚ.„¢[îÿ~	e2Ëﬁ?2Ï·¨îGêe–Bäi~§˝≥oΩ#ﬂ'ÔÏt–øiÍÌ¥*8¶”Lßÿ˛G≈hIVÕI>^≥áO2i‹U¯<ôˇ!ßÎº”ˇ√∏º¢›ŸlBkÚ˛‰&áˇE`˘SX|"˝ºÂü¡ÑUgëPË<P"C*Ïª≈›ˇ˜ãl:	«.…“xÎ=.›ÙUo Ú‰ıÑJ ˝Ç·Q9æ+µ·Uä=ùÊAÉ˘á°‹vîƒ#[£¸B1˝H:&çÖˆØ´q|U\ÁÁóÄÃó˝%
?ø+áìüQÙ±ˆ ÿ∞Ûº
d5ÜÂ\%ˆ·5EΩ{ót`xÚÆpv• ¥Õ>Â,¢ˆåÓ
˙ô ]Ø∫ÖêCèãk‡4eµAQèä∫ŒÊËáHﬁXOãK¶ÏzK}
g!™˛v&[DÖsBÁ´£∫ ˛≈ƒ!ˆ[{Q’‹	Rdmï4°U“ME‹Ÿe\‰‰EñüHÑUtˆ>Ñÿ qö≥¸zº⁄€<ôø€∫bÆ¸§êì» 7 ⁄ï$ô`¨∏Ó7 ´€È⁄›˛’  ˇˇ ûnN