import { Item } from './types';

export interface ItemGroup {
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

export const normalizeString = (str: string | null | undefined): string => 
  (str || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();

export const getSafeDocId = (name: string | null | undefined): string => {
  const normalized = normalizeString(name);
  return normalized.replace(/[^a-z0-9]/gi, '_');
};

export const SECTORS = [
  'CPSMS', 'CME', 'Clínica Geral', 'Higienização', 'Direção', 
  'Recepção', 'SAME', 'Copa', 'Administrativo', 'TI', 'Regulação'
];

export const SECTOR_COLORS: Record<string, string> = {
  'CPSMS': '#0284c7',
  'CME': '#7c3aed',
  'Clínica Geral': '#059669',
  'Higienização': '#6366f1',
  'Direção': '#ef4444',
  'Recepção': '#14b8a6',
  'SAME': '#7c2d12',
  'Copa': '#84cc16',
  'Administrativo': '#8b5cf6',
  'TI': '#1e293b',
  'Regulação': '#fb923c'
};

export const ROOMS = ['Sala A', 'Sala B', 'Almoxarifado Principal', 'Farmácia'];

export const CATEGORY_COLORS: Record<string, string> = {
  'Odontológico': '#0284c7',
  'Médico Hospitalar': '#ef4444',
  'Alimentício': '#f59e0b',
  'Expediente': '#3b82f6',
  'Higiene': '#10b981',
  'Radiológico': '#8b5cf6',
  'Saneante': '#06b6d4',
  'Copa & Cozinha': '#f97316',
  'Papelaria': '#0ea5e9',
  'EPI': '#ec4899',
  'Gráfica': '#fbbf24',
  'Informática': '#6366f1',
  'Limpeza': '#059669',
  'Anestésico': '#7c3aed',
  'Medicamentos': '#be123c',
  'Fisioterápicos': '#14b8a6',
  'Manutenção': '#57534e',
  'Outros': '#78716c',
};

export const getCategoryColor = (cat: string): string => {
  if (CATEGORY_COLORS[cat]) return CATEGORY_COLORS[cat];
  const hash = cat.split('').reduce((acc, char) => char.charCodeAt(0) + ((acc << 5) - acc), 0);
  return `hsl(${Math.abs(hash) % 360}, 70%, 50%)`;
};

export enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}
