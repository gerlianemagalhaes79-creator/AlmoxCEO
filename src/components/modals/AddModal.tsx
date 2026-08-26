import * as React from 'react';
import { X, Plus, Trash2, Copy, Save } from 'lucide-react';
import { motion } from 'motion/react';
import { ROOMS } from '../../constants';
import { Item } from '../../types';

export interface BulkEntryItem {
  id: string;
  name: string;
  initial_quantity: number;
  min_quantity: number;
  batch_number: string;
  expiry_date: string;
  is_indeterminate_expiry?: boolean;
  unit_price?: number;
  unit_measure?: string;
  medication_type?: string;
}

export interface BulkEntryData {
  supplier: string;
  category: string;
  origin: 'contract' | 'extra' | 'donation';
  room: string;
  items: BulkEntryItem[];
}

interface AddModalProps {
  showAddModal: boolean;
  setShowAddModal: (show: boolean) => void;
  handleAddItem: (e: React.FormEvent) => void;
  bulkEntry: BulkEntryData;
  setBulkEntry: React.Dispatch<React.SetStateAction<BulkEntryData>>;
  showNewCategoryInput: boolean;
  setShowNewCategoryInput: (show: boolean) => void;
  newCategoryName: string;
  setNewCategoryName: (name: string) => void;
  categories: string[];
  setCategories: React.Dispatch<React.SetStateAction<string[]>>;
  items: Item[];
  uniqueSuppliers: string[];
  updateBulkItem: (id: string, field: string, value: any) => void;
  duplicateBulkItem: (id: string) => void;
  removeBulkItemRow: (id: string) => void;
  addBulkItemRow: () => void;
}

export const AddModal: React.FC<AddModalProps> = ({
  showAddModal,
  setShowAddModal,
  handleAddItem,
  bulkEntry,
  setBulkEntry,
  showNewCategoryInput,
  setShowNewCategoryInput,
  newCategoryName,
  setNewCategoryName,
  categories,
  setCategories,
  items,
  uniqueSuppliers,
  updateBulkItem,
  duplicateBulkItem,
  removeBulkItemRow,
  addBulkItemRow,
}) => {
  if (!showAddModal) return null;

  return (
    <div className="fixed inset-0 bg-black/40 backdrop-blur-sm z-50 flex items-center justify-center p-6">
      <motion.div 
        initial={{ scale: 0.9, opacity: 0 }}
        animate={{ scale: 1, opacity: 1 }}
        className="bg-white w-full max-w-5xl rounded-[40px] p-10 shadow-2xl max-h-[90vh] overflow-y-auto"
      >
        <div className="flex justify-between items-center mb-8">
          <div>
            <h3 className="text-3xl font-black text-[#1C1917]">Entrada de Materiais</h3>
            <p className="text-[#78716C] font-medium">Cadastre múltiplos itens de uma vez</p>
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
                      className="px-4 py-3 bg-[#1C1917] text-white rounded-xl font-bold hover:bg-[#292524] transition-all"
                    >
                      +
                    </button>
                    <button 
                      type="button"
                      onClick={() => setShowNewCategoryInput(false)}
                      className="px-4 py-3 bg-[#F5F5F4] text-[#78716C] rounded-xl font-bold hover:bg-[#E7E5E4] transition-all"
                    >
                      <X size={18} />
                    </button>
                  </div>
                ) : (
                  <select 
                    className="flex-1 px-4 py-3 bg-white border border-[#E7E5E4] rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 font-bold"
                    value={bulkEntry.category}
                    onChange={e => {
                      if (e.target.value === '__NEW__') {
                        setShowNewCategoryInput(true);
                      } else {
                        setBulkEntry({...bulkEntry, category: e.target.value});
                      }
                    }}
                  >
                    {categories.map(cat => (
                      <option key={cat} value={cat}>{cat}</option>
                    ))}
                    <option value="__NEW__">+ Nova Categoria...</option>
                  </select>
                )}
              </div>
            </div>

            <div className="lg:col-span-1">
              <label className="block text-xs font-black text-[#78716C] uppercase tracking-widest mb-2">Origem</label>
              <select 
                className="w-full px-4 py-3 bg-white border border-[#E7E5E4] rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 font-bold"
                value={bulkEntry.origin}
                onChange={e => setBulkEntry({...bulkEntry, origin: e.target.value as 'contract' | 'extra' | 'donation'})}
              >
                <option value="extra">Extra</option>
                <option value="contract">Contrato</option>
                <option value="donation">Doação</option>
              </select>
            </div>

            <div className="lg:col-span-1">
              <label className="block text-xs font-black text-[#78716C] uppercase tracking-widest mb-2">Local de Armazenamento</label>
              <select 
                className="w-full px-4 py-3 bg-white border border-[#E7E5E4] rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 font-bold"
                value={bulkEntry.room}
                onChange={e => setBulkEntry({...bulkEntry, room: e.target.value})}
              >
                {ROOMS.map(r => (
                  <option key={r} value={r}>{r}</option>
                ))}
              </select>
            </div>
          </div>

          {/* Items List Section */}
          <div className="space-y-4">
            <div className="flex justify-between items-center">
              <h4 className="text-xl font-black text-[#1C1917]">Lista de Itens ({bulkEntry.items.length})</h4>
              <button 
                type="button"
                onClick={addBulkItemRow}
                className="px-4 py-2 bg-[#F5F5F4] hover:bg-[#E7E5E4] text-[#1C1917] font-bold rounded-xl flex items-center gap-2 transition-all"
              >
                <Plus size={18} /> Adicionar Linha
              </button>
            </div>

            <div className="overflow-x-auto rounded-2xl border border-[#E7E5E4]">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-[#FAFAF9] border-b border-[#E7E5E4] text-xs font-black text-[#78716C] uppercase tracking-wider">
                    <th className="py-3 px-4">Nome do Produto</th>
                    <th className="py-3 px-2 w-28">Emb/Medida</th>
                    {bulkEntry.category === 'Medicamentos' && (
                      <th className="py-3 px-2 w-32">Dosagem/Tipo</th>
                    )}
                    <th className="py-3 px-2 min-w-[100px] text-center">Qtd</th>
                    <th className="py-3 px-2 min-w-[100px] text-center">Estoque Mín</th>
                    <th className="py-3 px-2 w-24">Lote</th>
                    <th className="py-3 px-2 w-36">Validade</th>
                    <th className="py-3 px-2 text-center w-20">Ações</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#E7E5E4] bg-white">
                  {bulkEntry.items.map((item) => (
                    <tr key={item.id} className="hover:bg-[#FAFAF9]/50 transition-colors">
                      <td className="p-2">
                        <input 
                          required
                          list="product-suggestions"
                          type="text" 
                          placeholder="Nome do produto"
                          className="w-full px-3 py-2 bg-[#FAFAF9] border border-[#E7E5E4] rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 text-sm font-bold uppercase"
                          value={item.name}
                          onChange={e => updateBulkItem(item.id, 'name', e.target.value)}
                        />
                      </td>
                      <td className="px-2 w-28">
                        <select 
                          className="w-full px-2 py-2 bg-[#FAFAF9] border border-[#E7E5E4] rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 text-xs font-bold"
                          value={item.unit_measure || 'Unidade (UN)'}
                          onChange={e => updateBulkItem(item.id, 'unit_measure', e.target.value)}
                        >
                          <option value="Unidade (UN)">UN</option>
                          <option value="Caixa (CX)">CX</option>
                          <option value="Pacote (PCT)">PCT</option>
                          <option value="Frasco (FR)">FR</option>
                          <option value="Ampola (AMP)">AMP</option>
                          <option value="Bisnaga (BSG)">BSG</option>
                          <option value="Envelope (ENV)">ENV</option>
                          <option value="Litro (L)">L</option>
                          <option value="Quilo (KG)">KG</option>
                          <option value="Metro (M)">M</option>
                          <option value="Rolo (RL)">RL</option>
                          <option value="Resma">Resma</option>
                          <option value="Par (PR)">PR</option>
                          <option value="Outro">Outro</option>
                        </select>
                      </td>
                      {bulkEntry.category === 'Medicamentos' && (
                        <td className="px-2 w-32">
                          <input 
                            type="text" 
                            placeholder="Ex: 500mg, CP"
                            className="w-full px-3 py-2 bg-[#FAFAF9] border border-[#E7E5E4] rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 text-xs font-bold uppercase"
                            value={item.medication_type || ''}
                            onChange={e => updateBulkItem(item.id, 'medication_type', e.target.value.toUpperCase())}
                          />
                        </td>
                      )}
                      <td className="px-2 min-w-[100px]">
                        <input 
                          required
                          type="number" 
                          min="1"
                          placeholder="Qtd"
                          className="w-full px-3 py-2 bg-emerald-50 border-2 border-emerald-300 rounded-xl focus:ring-2 focus:ring-emerald-500 text-sm text-emerald-950 font-black shadow-sm text-center"
                          value={isNaN(item.initial_quantity) ? '' : item.initial_quantity}
                          onChange={e => updateBulkItem(item.id, 'initial_quantity', e.target.value === '' ? NaN : parseInt(e.target.value))}
                        />
                      </td>
                      <td className="px-2 min-w-[100px]">
                        <input 
                          required
                          type="number" 
                          min="0"
                          placeholder="Mín"
                          className="w-full px-3 py-2 bg-amber-50 border border-amber-300 rounded-xl focus:ring-2 focus:ring-amber-500 text-sm text-amber-950 font-extrabold text-center"
                          value={isNaN(item.min_quantity) ? '' : item.min_quantity}
                          onChange={e => updateBulkItem(item.id, 'min_quantity', e.target.value === '' ? NaN : parseInt(e.target.value))}
                        />
                      </td>
                      <td className="px-2 w-24">
                        <input 
                          type="text" 
                          placeholder="Lote"
                          className="w-full px-3 py-2 bg-[#FAFAF9] border border-[#E7E5E4] rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 text-xs font-bold uppercase"
                          value={item.batch_number}
                          onChange={e => updateBulkItem(item.id, 'batch_number', e.target.value.toUpperCase())}
                        />
                      </td>
                      <td className="px-2 w-36">
                        <div className="space-y-1">
                          <input 
                            type="date" 
                            disabled={item.is_indeterminate_expiry}
                            className={`w-full px-2 py-2 bg-[#FAFAF9] border border-[#E7E5E4] rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 text-xs font-bold ${item.is_indeterminate_expiry ? 'opacity-40 bg-gray-100 cursor-not-allowed' : ''}`}
                            value={item.is_indeterminate_expiry ? '' : item.expiry_date}
                            onChange={e => updateBulkItem(item.id, 'expiry_date', e.target.value)}
                          />
                          <label className="flex items-center gap-1.5 cursor-pointer text-[10px] text-gray-500 hover:text-gray-700 select-none">
                            <input 
                              type="checkbox"
                              checked={!!item.is_indeterminate_expiry}
                              onChange={e => {
                                updateBulkItem(item.id, 'is_indeterminate_expiry', e.target.checked);
                                if (e.target.checked) {
                                  updateBulkItem(item.id, 'expiry_date', '');
                                }
                              }}
                              className="rounded border-gray-300 text-blue-600 focus:ring-blue-500 w-3 h-3"
                            />
                            <span>Indeterminada</span>
                          </label>
                        </div>
                      </td>
                      <td className="p-2 text-center">
                        <div className="flex items-center justify-center gap-1">
                          <button 
                            type="button"
                            title="Duplicar item (mesmo produto/fornecedor com novo lote)"
                            onClick={() => duplicateBulkItem(item.id)}
                            className="p-2 text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                          >
                            <Copy size={16} />
                          </button>
                          <button 
                            type="button"
                            disabled={bulkEntry.items.length === 1}
                            onClick={() => removeBulkItemRow(item.id)}
                            className="p-2 text-[#EF4444] hover:bg-[#FEF2F2] rounded-lg transition-colors disabled:opacity-30 disabled:hover:bg-transparent"
                          >
                            <Trash2 size={16} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {/* Quick Suggestions for Medication Types */}
            {bulkEntry.category === 'Medicamentos' && (
              <div className="bg-blue-50/60 border border-blue-200/80 rounded-2xl p-4 flex flex-wrap items-center gap-2">
                <span className="text-xs font-bold text-blue-900">Sugestões de Tipo de Medicamento:</span>
                {['COMPRIMIDO (CP)', 'AMPOLA (AMP)', 'FRASCO (FR)', 'GOTAS (GOTA)', 'POMADA', 'SPRAY', 'PORTARIA 344/98 (CONTROLADO)'].map((tag) => (
                  <button
                    key={tag}
                    type="button"
                    onClick={() => {
                      const lastItem = bulkEntry.items[bulkEntry.items.length - 1];
                      if (lastItem) {
                        updateBulkItem(lastItem.id, 'medication_type', tag);
                      }
                    }}
                    className="text-[11px] font-semibold bg-white border border-blue-300 text-blue-700 px-2.5 py-1 rounded-lg hover:bg-blue-600 hover:text-white transition-colors"
                  >
                    + {tag}
                  </button>
                ))}
              </div>
            )}
          </div>

          <datalist id="product-suggestions">
            {Array.from(new Set(items.map(i => i.name))).map(name => (
              <option key={name} value={name} />
            ))}
          </datalist>

          <datalist id="supplier-suggestions">
            {uniqueSuppliers.map(s => (
              <option key={s} value={s} />
            ))}
          </datalist>

          <div className="flex gap-4 pt-4 border-t border-[#E7E5E4]">
            <button 
              type="button"
              onClick={() => setShowAddModal(false)}
              className="flex-1 px-6 py-4 rounded-2xl font-bold text-[#78716C] hover:bg-[#F5F5F4] transition-all"
            >
              Cancelar
            </button>
            <button 
              type="submit"
              className="flex-[2] px-6 py-4 bg-[#1C1917] text-white rounded-2xl font-bold hover:bg-[#292524] transition-all shadow-lg shadow-[#1C1917]/20 flex items-center justify-center gap-3"
            >
              <Save size={20} /> Finalizar Entrada de {bulkEntry.items.length} Itens
            </button>
          </div>
        </form>
      </motion.div>
    </div>
  );
};
