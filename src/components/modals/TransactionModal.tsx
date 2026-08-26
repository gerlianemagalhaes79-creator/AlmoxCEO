import * as React from 'react';
import { X, Search, Trash2, Upload } from 'lucide-react';
import { motion } from 'motion/react';
import { Item } from '../../types';
import { SECTORS, normalizeString } from '../../constants';

interface TransactionModalProps {
  showTransactionModal: {
    show: boolean;
    type: 'entry' | 'exit';
    item?: Item | null;
  };
  setShowTransactionModal: (modal: { show: boolean; type: 'entry' | 'exit'; item?: Item | null }) => void;
  handleTransaction: (e: React.FormEvent) => void;
  selectedItemId: string;
  setSelectedItemId: (id: string) => void;
  items: Item[];
  weeklyExitRates: Record<string, number>;
  basket: Array<{ item_id: string; quantity: number; exitReason?: string }>;
  setBasket: React.Dispatch<React.SetStateAction<Array<{ item_id: string; quantity: number; exitReason?: string }>>>;
  transactionQty: number;
  setTransactionQty: (qty: number) => void;
  transactionMinStock: number;
  setTransactionMinStock: (min: number) => void;
  modalSector: string;
  setModalSector: (sec: string) => void;
  exitReason: 'consumo' | 'doacao' | 'vencido' | 'perda';
  setExitReason: (reason: 'consumo' | 'doacao' | 'vencido' | 'perda') => void;
  expiryReason: string;
  setExpiryReason: (reason: string) => void;
  donationUnitName: string;
  setDonationUnitName: (name: string) => void;
  donationUnitAddress: string;
  setDonationUnitAddress: (addr: string) => void;
  donationUnitCNPJ: string;
  setDonationUnitCNPJ: (cnpj: string) => void;
  donationRevisionDate: string;
  setDonationRevisionDate: (date: string) => void;
  letterheadImage: string | null;
  setLetterheadImage: (img: string | null) => void;
  modalSearchTerm: string;
  setModalSearchTerm: (term: string) => void;
  selectedItemName: string;
  setSelectedItemName: (name: string) => void;
  inventoryLocation: 'Almoxarifado' | 'Farmácia';
  isNearExpiry: (item: Item) => boolean;
  showToast: (msg: string, type: 'success' | 'error' | 'info') => void;
}

export const TransactionModal: React.FC<TransactionModalProps> = ({
  showTransactionModal,
  setShowTransactionModal,
  handleTransaction,
  selectedItemId,
  setSelectedItemId,
  items,
  weeklyExitRates,
  basket,
  setBasket,
  transactionQty,
  setTransactionQty,
  transactionMinStock,
  setTransactionMinStock,
  modalSector,
  setModalSector,
  exitReason,
  setExitReason,
  expiryReason,
  setExpiryReason,
  donationUnitName,
  setDonationUnitName,
  donationUnitAddress,
  setDonationUnitAddress,
  donationUnitCNPJ,
  setDonationUnitCNPJ,
  donationRevisionDate,
  setDonationRevisionDate,
  letterheadImage,
  setLetterheadImage,
  modalSearchTerm,
  setModalSearchTerm,
  selectedItemName,
  setSelectedItemName,
  inventoryLocation,
  isNearExpiry,
  showToast,
}) => {
  if (!showTransactionModal.show) return null;

  return (
        <div className="fixed inset-0 bg-black/40 backdrop-blur-sm z-50 flex items-center justify-center p-6">
          <motion.div 
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-white w-full max-w-lg rounded-3xl p-8 shadow-2xl max-h-[90vh] overflow-y-auto"
          >
            <h3 className="text-2xl font-bold mb-6">
              {showTransactionModal.type === 'entry' ? 'Registrar Entrada' : 'Registrar Saída'}
            </h3>
            
            <form onSubmit={handleTransaction} className="space-y-6">
              {showTransactionModal.type === 'entry' ? (
                <>
                  {showTransactionModal.item ? (
                    <div className="mb-6">
                      <p className="text-[#78716C] font-medium">{showTransactionModal.item.name}</p>
                      <p className="text-xs font-bold text-emerald-600 mt-1">
                        Disponível em estoque: {showTransactionModal.item.quantity} unidades
                      </p>
                    </div>
                  ) : (
                    <div className="mb-6">
                      <label className="block text-sm font-bold text-[#57534E] mb-2">Selecionar Item</label>
                      <select 
                        required
                        className="w-full px-4 py-3 bg-[#F5F5F4] border-none rounded-xl focus:ring-2 focus:ring-[#1C1917]/10"
                        value={selectedItemId}
                        onChange={e => setSelectedItemId(e.target.value)}
                      >
                        <option value="">Selecione um item...</option>
                        {items.map(item => (
                          <option key={item.id} value={item.id}>
                            {item.name} (Lote: {item.batch_number || 'N/A'}) - {item.quantity} un.
                          </option>
                        ))}
                      </select>
                    </div>
                  )}
                  <div>
                    <label className="block text-sm font-extrabold text-[#57534E] mb-2 text-center uppercase tracking-wider">Quantidade a Adicionar</label>
                    <div className="flex items-center justify-center gap-4 py-2">
                      <button 
                        type="button"
                        onClick={() => setTransactionQty(Math.max(1, transactionQty - 1))}
                        className="w-12 h-12 rounded-2xl bg-slate-100 hover:bg-slate-200 border border-slate-300 flex items-center justify-center text-2xl font-black text-slate-800 shadow-sm transition-all"
                      >
                        -
                      </button>
                      <input 
                        type="number"
                        min="1"
                        value={transactionQty}
                        onChange={e => setTransactionQty(Math.max(1, parseInt(e.target.value) || 0))}
                        className="text-3xl font-black w-32 py-2 px-3 text-center bg-emerald-50 text-emerald-950 border-2 border-emerald-500 rounded-2xl shadow-inner focus:ring-2 focus:ring-emerald-500/30"
                      />
                      <button 
                        type="button"
                        onClick={() => setTransactionQty(transactionQty + 1)}
                        className="w-12 h-12 rounded-2xl bg-slate-100 hover:bg-slate-200 border border-slate-300 flex items-center justify-center text-2xl font-black text-slate-800 shadow-sm transition-all"
                      >
                        +
                      </button>
                    </div>
                  </div>

                  <div>
                    <label className="block text-sm font-bold text-[#57534E] mb-2">Estoque Mínimo (5 Semanas)</label>
                    <input 
                      type="number"
                      placeholder="Calculando..."
                      className="w-full px-4 py-3 bg-[#F5F5F4] border-none rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 font-bold"
                      value={isNaN(transactionMinStock) ? (
                        (() => {
                          const item = showTransactionModal.item || items.find(i => i.id === selectedItemId);
                          if (item) {
                            const weeklyRate = weeklyExitRates[item.name] || 0;
                            return weeklyRate > 0 ? Math.ceil(weeklyRate * 5) : item.min_quantity;
                          }
                          return '';
                        })()
                      ) : transactionMinStock}
                      onChange={e => setTransactionMinStock(parseInt(e.target.value))}
                    />
                    <p className="text-[10px] text-[#A8A29E] mt-1 font-medium italic">
                      Deixe em branco para usar o cálculo automático do sistema.
                    </p>
                  </div>
                </>
              ) : (
                <div className="space-y-6">
                  <div>
                    <label className="block text-sm font-bold text-[#57534E] mb-2">Motivo da Saída</label>
                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                      <button
                        type="button"
                        onClick={() => {
                          setExitReason('consumo');
                          setModalSector(SECTORS[0]);
                        }}
                        className={`px-3 py-2 rounded-xl text-xs font-bold border transition-all ${exitReason === 'consumo' ? 'bg-[#1C1917] text-white border-[#1C1917]' : 'bg-white text-[#78716C] border-[#E7E5E4] hover:bg-[#F5F5F4]'}`}
                      >
                        Consumo
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setExitReason('doacao');
                          setModalSector('');
                        }}
                        className={`px-3 py-2 rounded-xl text-xs font-bold border transition-all ${exitReason === 'doacao' ? 'bg-[#1C1917] text-white border-[#1C1917]' : 'bg-white text-[#78716C] border-[#E7E5E4] hover:bg-[#F5F5F4]'}`}
                      >
                        Doação
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setExitReason('vencido');
                          setModalSector('Descarte/Vencimento');
                          setExpiryReason('');
                        }}
                        className={`px-3 py-2 rounded-xl text-xs font-bold border transition-all ${exitReason === 'vencido' ? 'bg-[#1C1917] text-white border-[#1C1917]' : 'bg-white text-[#78716C] border-[#E7E5E4] hover:bg-[#F5F5F4]'}`}
                      >
                        Vencido
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setExitReason('perda');
                          setModalSector('Perda/Avaria');
                          setExpiryReason('');
                        }}
                        className={`px-3 py-2 rounded-xl text-xs font-bold border transition-all ${exitReason === 'perda' ? 'bg-[#1C1917] text-white border-[#1C1917]' : 'bg-white text-[#78716C] border-[#E7E5E4] hover:bg-[#F5F5F4]'}`}
                      >
                        Perda/Avaria
                      </button>
                    </div>
                  </div>

                  <div>
                    <label className="block text-sm font-bold text-[#57534E] mb-2">
                      {exitReason === 'doacao' ? 'Destinatário da Doação' : 
                       (exitReason === 'vencido' || exitReason === 'perda') ? 'Classificação' : 'Setor de Destino'}
                    </label>
                    {exitReason === 'doacao' ? (
                      <div className="space-y-4">
                        <div>
                          <label className="block text-[10px] font-bold text-[#A8A29E] uppercase mb-1 ml-1">Unidade Doadora</label>
                          <input 
                            required
                            type="text"
                            placeholder="CEO - Centro de Especialidades Odontológicas"
                            className="w-full px-4 py-3 bg-[#F5F5F4] border-none rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 font-bold"
                            value={donationUnitName || 'CEO - Centro de Especialidades Odontológicas'}
                            onChange={e => setDonationUnitName(e.target.value)}
                          />
                        </div>
                        <div>
                          <label className="block text-[10px] font-bold text-[#A8A29E] uppercase mb-1 ml-1">Unidade Receptora (Nome)</label>
                          <input 
                            required
                            type="text"
                            placeholder="Nome da unidade receptora..."
                            className="w-full px-4 py-3 bg-[#F5F5F4] border-none rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 font-bold"
                            value={modalSector}
                            onChange={e => setModalSector(e.target.value)}
                          />
                        </div>
                        <div className="grid grid-cols-2 gap-3">
                          <div>
                            <label className="block text-[10px] font-bold text-[#A8A29E] uppercase mb-1 ml-1">Endereço Receptora</label>
                            <input 
                              required
                              type="text"
                              placeholder="Endereço..."
                              className="w-full px-4 py-3 bg-[#F5F5F4] border-none rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 font-bold text-xs"
                              value={donationUnitAddress}
                              onChange={e => setDonationUnitAddress(e.target.value)}
                            />
                          </div>
                          <div>
                            <label className="block text-[10px] font-bold text-[#A8A29E] uppercase mb-1 ml-1">CNPJ Receptora</label>
                            <input 
                              required
                              type="text"
                              placeholder="00.000.000/0000-00"
                              className="w-full px-4 py-3 bg-[#F5F5F4] border-none rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 font-bold text-xs"
                              value={donationUnitCNPJ}
                              onChange={e => setDonationUnitCNPJ(e.target.value)}
                            />
                          </div>
                        </div>
                        <div>
                          <label className="block text-[10px] font-bold text-[#A8A29E] uppercase mb-1 ml-1">Papel Timbrado (Opcional - JPEG/PNG)</label>
                          <div className="flex items-center gap-3">
                            <label className="flex-1 cursor-pointer group">
                              <div className="flex items-center gap-2 px-4 py-3 bg-[#F5F5F4] border-2 border-dashed border-[#E7E5E4] rounded-xl hover:border-[#1C1917]/20 transition-all">
                                <Upload size={16} className="text-[#A8A29E] group-hover:text-[#1C1917]" />
                                <span className="text-xs font-bold text-[#78716C] group-hover:text-[#1C1917]">
                                  {letterheadImage ? 'Alterar Imagem' : 'Selecionar Timbrado'}
                                </span>
                              </div>
                              <input 
                                type="file" 
                                accept="image/*" 
                                className="hidden" 
                                onChange={(e) => {
                                  const file = e.target.files?.[0];
                                  if (file) {
                                    const reader = new FileReader();
                                    reader.onloadend = () => {
                                      setLetterheadImage(reader.result as string);
                                    };
                                    reader.readAsDataURL(file);
                                  }
                                }}
                              />
                            </label>
                            {letterheadImage && (
                              <button 
                                type="button"
                                onClick={() => setLetterheadImage(null)}
                                className="p-3 bg-red-50 text-red-500 rounded-xl hover:bg-red-100 transition-colors"
                                title="Remover imagem"
                              >
                                <Trash2 size={16} />
                              </button>
                            )}
                          </div>
                          {letterheadImage && (
                            <div className="mt-2 relative w-full h-12 bg-white rounded-lg border border-[#E7E5E4] overflow-hidden">
                              <img 
                                src={letterheadImage} 
                                alt="Preview" 
                                className="w-full h-full object-contain"
                              />
                            </div>
                          )}
                        </div>
                        <div>
                          <label className="block text-[10px] font-bold text-[#A8A29E] uppercase mb-1 ml-1">Data da Última Revisão</label>
                          <input 
                            required
                            type="text"
                            placeholder="Ex: 24/04/2026"
                            className="w-full px-4 py-3 bg-[#F5F5F4] border-none rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 font-bold"
                            value={donationRevisionDate}
                            onChange={e => setDonationRevisionDate(e.target.value)}
                          />
                        </div>
                      </div>
                    ) : (exitReason === 'vencido' || exitReason === 'perda') ? (
                      <div className="bg-rose-50 p-4 rounded-2xl border border-rose-100">
                        <p className="text-xs font-bold text-rose-700 uppercase tracking-widest mb-1">Descarte por {exitReason === 'vencido' ? 'Vencimento' : 'Perda/Avaria'}</p>
                        <p className="text-sm text-rose-600">Esta movimentação será registrada como {modalSector}.</p>
                      </div>
                    ) : (
                      <select 
                        required
                        className="w-full px-4 py-3 bg-[#F5F5F4] border-none rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 font-bold"
                        value={modalSector}
                        onChange={e => setModalSector(e.target.value)}
                      >
                        <option value="">Selecione o setor de destino...</option>
                        <option value="Farmácia (Consumo Interno)">Farmácia (Consumo Interno)</option>
                        {SECTORS.map(sector => (
                          <option key={sector} value={sector}>{sector}</option>
                        ))}
                      </select>
                    )}
                  </div>

                  {(exitReason === 'vencido' || exitReason === 'perda') && (
                    <motion.div 
                      initial={{ opacity: 0, y: -10 }}
                      animate={{ opacity: 1, y: 0 }}
                      className="space-y-2"
                    >
                      <label className="block text-sm font-bold text-[#57534E]">Justificativa do {exitReason === 'vencido' ? 'Vencimento' : 'Descarte'}</label>
                      <textarea 
                        required
                        placeholder={exitReason === 'vencido' ? "Explique por que o item venceu no estoque..." : "Explique o motivo da perda ou avaria..."}
                        className="w-full px-4 py-3 bg-[#F5F5F4] border-none rounded-xl focus:ring-2 focus:ring-[#1C1917]/10 font-bold text-sm min-h-[100px] resize-none"
                        value={expiryReason}
                        onChange={e => setExpiryReason(e.target.value)}
                      />
                    </motion.div>
                  )}

                  <div className="space-y-4">
                    <label className="block text-sm font-bold text-[#57534E]">Itens para Saída</label>
                    {basket.map((b, index) => {
                      const item = items.find(i => i.id === b.item_id);
                      return (
                        <div key={index} className="flex items-center gap-4 bg-[#F5F5F4] p-4 rounded-2xl">
                          <div className="flex-1">
                            <p className="font-bold text-sm">{item?.name || 'Item não encontrado'}</p>
                            <p className="text-[10px] text-[#78716C]">Lote: {item?.batch_number || 'N/A'} | Estoque: {item?.quantity || 0}</p>
                          </div>
                          <div className="flex items-center gap-3">
                            <button 
                              type="button"
                              onClick={() => {
                                const newBasket = [...basket];
                                newBasket[index].quantity = Math.max(1, newBasket[index].quantity - 1);
                                setBasket(newBasket);
                              }}
                              className="w-8 h-8 rounded-lg bg-white flex items-center justify-center font-bold hover:bg-gray-100"
                            >
                              -
                            </button>
                            <input 
                              type="number"
                              min="1"
                              max={item?.quantity || 999}
                              value={b.quantity}
                              onChange={e => {
                                const val = Math.max(1, Math.min(item?.quantity || 999, parseInt(e.target.value) || 0));
                                const newBasket = [...basket];
                                newBasket[index].quantity = val;
                                setBasket(newBasket);
                              }}
                              className="font-bold w-16 text-center bg-transparent border-none focus:ring-0 text-sm"
                            />
                            <button 
                              type="button"
                              onClick={() => {
                                const newBasket = [...basket];
                                newBasket[index].quantity = Math.min(item?.quantity || 999, newBasket[index].quantity + 1);
                                setBasket(newBasket);
                              }}
                              className="w-8 h-8 rounded-lg bg-white flex items-center justify-center font-bold hover:bg-gray-100"
                            >
                              +
                            </button>
                            <button 
                              type="button"
                              onClick={() => setBasket(basket.filter((_, i) => i !== index))}
                              className="text-rose-500 hover:text-rose-700 ml-2"
                            >
                              <X size={18} />
                            </button>
                          </div>
                        </div>
                      );
                    })}

                    <div className="space-y-4">
                      <div className="flex flex-col gap-4">
                        <div className="flex-1">
                          <label className="block text-[10px] font-bold text-[#A8A29E] uppercase mb-1 ml-1">1. Escolha o Item</label>
                          <div className="relative">
                            <div className="relative">
                              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-[#A8A29E]" size={16} />
                              <input 
                                autoFocus
                                type="text" 
                                placeholder="Pesquisar item..."
                                className="w-full pl-10 pr-4 py-3 bg-[#F5F5F4] border border-[#E7E5E4] rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-[#1C1917]/10 font-bold"
                                value={modalSearchTerm}
                                onChange={(e) => {
                                  setModalSearchTerm(e.target.value);
                                  if (selectedItemName) setSelectedItemName('');
                                }}
                                onKeyDown={(e) => {
                                  if (e.key === 'Enter' && !modalSearchTerm && basket.length > 0) {
                                    e.preventDefault();
                                    // Submit the form
                                    const form = e.currentTarget.closest('form');
                                    if (form) form.requestSubmit();
                                  }
                                }}
                              />
                            </div>

                            {modalSearchTerm.length >= 2 && !selectedItemName && (
                              <motion.div 
                                initial={{ opacity: 0, y: -5 }}
                                animate={{ opacity: 1, y: 0 }}
                                className="absolute left-0 right-0 top-full mt-1 bg-white border border-[#E7E5E4] rounded-xl shadow-lg z-50 max-h-48 overflow-y-auto"
                              >
                                {(items.filter(i => i.quantity > 0) as Item[])
                                  .filter(item => {
                                    const combined = `${item.name} ${item.batch_number || ''}`;
                                    return normalizeString(combined).includes(normalizeString(modalSearchTerm));
                                  })
                                  .sort((a, b) => a.name.localeCompare(b.name))
                                  .slice(0, 10)
                                  .map(item => (
                                    <button
                                      key={item.id}
                                      type="button"
                                      onClick={() => {
                                        if (basket.some(b => b.item_id === item.id)) {
                                          showToast('Este lote já está na lista de saída.', 'error');
                                          return;
                                        }
                                        setBasket([...basket, { item_id: item.id, quantity: 1 }]);
                                        setModalSearchTerm('');
                                        setSelectedItemName('');
                                      }}
                                      className="w-full px-4 py-3 text-left hover:bg-[#F5F5F4] transition-all border-b border-[#F5F5F4] last:border-none flex justify-between items-center"
                                    >
                                      <div>
                                        <p className="font-bold text-sm text-[#1C1917]">{item.name}</p>
                                        <p className="text-[10px] text-[#78716C] font-mono">Lote: {item.batch_number || '---'}</p>
                                      </div>
                                      <div className="flex flex-col items-end">
                                        <span className="text-[10px] font-black text-emerald-600 bg-emerald-50 px-1.5 py-0.5 rounded uppercase">
                                          {item.quantity} un.
                                        </span>
                                        {item.expiry_date && (
                                          <span className={`text-[8px] font-bold ${isNearExpiry(item) ? 'text-rose-600' : 'text-[#A8A29E]'}`}>
                                            {item.expiry_date === 'Indeterminada' ? 'Indeterminada' : new Date(item.expiry_date).toLocaleDateString('pt-BR')}
                                          </span>
                                        )}
                                      </div>
                                    </button>
                                  ))
                                }
                              </motion.div>
                            )}
                          </div>
                        </div>

                        {selectedItemName && (
                          <div className="flex-1">
                            <label className="block text-[10px] font-bold text-[#A8A29E] uppercase mb-1 ml-1">2. Escolha o Lote</label>
                            <select 
                              className="w-full px-4 py-3 bg-[#F5F5F4] border-none rounded-xl text-sm focus:ring-2 focus:ring-[#1C1917]/10"
                              value={selectedItemId}
                              onChange={e => {
                                const id = e.target.value;
                                if (!id) return;
                                if (basket.some(b => b.item_id === id)) {
                                  alert('Este lote já está na lista de saída.');
                                  return;
                                }
                                setBasket([...basket, { item_id: id, quantity: 1 }]);
                                setSelectedItemId('');
                                setSelectedItemName('');
                                setModalSearchTerm('');
                              }}
                            >
                              <option value="">Selecione o lote...</option>
                              {items
                                .filter(i => i.name === selectedItemName && i.quantity > 0 && !basket.some(b => b.item_id === i.id))
                                .map(item => (
                                  <option key={item.id} value={item.id}>
                                    Lote: {item.batch_number || 'S/N'} ({item.quantity} un.) {item.expiry_date ? `- Venc: ${new Date(item.expiry_date).toLocaleDateString('pt-BR')}` : ''}
                                  </option>
                                ))
                              }
                            </select>
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                </div>
              )}
              
              <div className="flex gap-3 pt-4">
                <button 
                  type="button"
                  onClick={() => {
                    setShowTransactionModal({ show: false, type: 'entry' });
                    setLetterheadImage(null);
                  }}
                  className="flex-1 px-4 py-3 rounded-xl font-bold text-[#78716C] hover:bg-[#F5F5F4] transition-all"
                >
                  Cancelar
                </button>
                <button 
                  type="submit"
                  disabled={showTransactionModal.type === 'exit' && basket.length === 0}
                  className={`flex-1 px-4 py-3 text-white rounded-xl font-bold transition-all disabled:opacity-50 ${showTransactionModal.type === 'entry' ? 'bg-emerald-600 hover:bg-emerald-700' : 'bg-rose-600 hover:bg-rose-700'}`}
                >
                  Confirmar {showTransactionModal.type === 'exit' && basket.length > 0 && `(${basket.length})`}
                </button>
              </div>
            </form>
          </motion.div>
        </div>
  );
};
