import * as React from 'react';
import { RotateCcw, AlertTriangle, CheckCircle, Clock, X, Package } from 'lucide-react';
import { motion } from 'motion/react';
import { Item, MaterialRequest, RequestItem } from '../../types';

interface DevolutionTabProps {
  selectedSector: string;
  devolutionSubTab: 'my_returns' | 'eligible_deliveries' | 'sector_stock';
  setDevolutionSubTab: (tab: 'my_returns' | 'eligible_deliveries' | 'sector_stock') => void;
  requests: MaterialRequest[];
  allRequestItems: RequestItem[];
  items: Item[];
  isExpired: (item: Item) => boolean;
  isNearExpiry: (item: Item) => boolean;
  setDevolutionBasket: React.Dispatch<React.SetStateAction<any[]>>;
  setSelectedDevProduct?: (p: string) => void;
  setDevolutionReason: (reason: string) => void;
  setDevolutionObservation: (obs: string) => void;
  setShowDevolutionModal: (modal: { show: boolean; request?: MaterialRequest }) => void;
  setShowRequestDetailModal: (modal: { show: boolean; request?: MaterialRequest }) => void;
  showToast?: (msg: string, type: 'success' | 'error' | 'info') => void;
}

export const DevolutionTab: React.FC<DevolutionTabProps> = ({
  selectedSector,
  devolutionSubTab,
  setDevolutionSubTab,
  requests,
  allRequestItems,
  items,
  isExpired,
  isNearExpiry,
  setDevolutionBasket,
  setSelectedDevProduct,
  setDevolutionReason,
  setDevolutionObservation,
  setShowDevolutionModal,
  setShowRequestDetailModal,
  showToast,
}) => {
  return (
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
                      Devolução de Materiais
                    </h2>
                  </div>
                  <p className="text-xs sm:text-sm text-slate-600 font-medium max-w-xl">
                    Gerencie o retorno de materiais do setor <strong className="text-amber-700">{selectedSector}</strong> ao almoxarifado de forma simples e organizada.
                  </p>
                </div>
                <button
                  onClick={() => {
                    setDevolutionBasket([]);
                    if (setSelectedDevProduct) setSelectedDevProduct('');
                    setDevolutionReason('Não teve uso');
                    setDevolutionObservation('');
                    setShowDevolutionModal({ show: true });
                  }}
                  className="z-10 bg-gradient-to-r from-amber-600 to-amber-700 text-white px-6 py-3.5 rounded-2xl text-xs font-black uppercase tracking-wider hover:from-amber-700 hover:to-amber-800 transition-all shadow-md shadow-amber-600/20 flex items-center justify-center gap-2 whitespace-nowrap self-start md:self-auto hover:-translate-y-0.5 active:translate-y-0"
                >
                  <RotateCcw size={16} />
                  Solicitar Devolução
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
                  Minhas Devoluções
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
                  Entregas Elegíveis para Devolução
                </button>
              </div>

              {/* Subtab Content */}
              {devolutionSubTab === 'my_returns' && (
                <div className="bg-white rounded-3xl border border-slate-200/80 shadow-xs overflow-hidden">
                  <div className="p-6 border-b border-slate-100">
                    <h3 className="text-base font-black text-slate-900">Histórico de Solicitações de Devolução</h3>
                    <p className="text-xs text-slate-500 font-medium">Acompanhe o andamento e o parecer das solicitações de devolução do seu setor.</p>
                  </div>

                  <div className="p-4 sm:p-6 space-y-3">
                    {requests
                      .filter(r => r.sector === selectedSector && r.isReturn && !r.deletedAt)
                      .map(req => {
                        const reqItems = allRequestItems.filter(ri => ri.request_id === req.id);
                        return (
                          <div 
                            key={req.id} 
                            className="bg-slate-50/60 border border-slate-200/60 hover:border-amber-200 rounded-2xl p-4 sm:p-5 transition-all hover:bg-amber-50/10 flex flex-col sm:flex-row sm:items-center justify-between gap-4"
                          >
                            <div className="space-y-2 flex-1">
                              <div className="flex flex-wrap items-center gap-2">
                                <span className="font-black text-xs text-slate-900 bg-white px-2.5 py-1 rounded-lg border border-slate-200">
                                  #{req.id.slice(-5).toUpperCase()}
                                </span>
                                <span className="text-xs text-slate-500 font-semibold">
                                  - {new Date(req.date).toLocaleDateString('pt-BR')}
                                </span>
                                <span className={`text-[10px] font-black px-3 py-1 rounded-full uppercase tracking-wider border ${
                                  req.status === 'DEVOLUCAO_PENDENTE' ? 'bg-amber-50 text-amber-700 border-amber-200' :
                                  req.status === 'DEVOLUCAO_APROVADA' ? 'bg-emerald-50 text-emerald-700 border-emerald-200' :
                                  req.status === 'DEVOLUCAO_RECUSADA' ? 'bg-rose-50 text-rose-700 border-rose-200' :
                                  'bg-slate-100 text-slate-700 border-slate-200'
                                }`}>
                                  {req.status === 'DEVOLUCAO_PENDENTE' ? 'PENDENTE' :
                                   req.status === 'DEVOLUCAO_APROVADA' ? 'APROVADO' :
                                   req.status === 'DEVOLUCAO_RECUSADA' ? 'RECUSADO' :
                                   req.status}
                                </span>
                                {req.returnReason && (
                                  <span className="text-[10px] font-bold text-amber-800 bg-amber-100/80 px-2.5 py-0.5 rounded-md">
                                    Motivo: {req.returnReason}
                                  </span>
                                )}
                              </div>

                              <div className="flex flex-wrap gap-1.5 pt-1">
                                {reqItems.map(item => (
                                  <span key={item.id} className="text-xs font-bold px-2.5 py-1 rounded-lg bg-white border border-slate-200 text-slate-800 shadow-2xs">
                                    {item.product_name} <strong className="text-amber-700">({item.quantity_requested})</strong>
                                  </span>
                                ))}
                              </div>

                              {req.observation && (
                                <p className="text-xs text-slate-500 italic bg-white/60 p-2 rounded-xl border border-slate-100">
                                  "{req.observation}"
                                </p>
                              )}
                            </div>

                            <button 
                              onClick={() => setShowRequestDetailModal({ show: true, request: req })}
                              className="bg-slate-900 text-white px-4 py-2.5 rounded-xl text-xs font-bold hover:bg-slate-800 transition-all shadow-xs self-end sm:self-center whitespace-nowrap"
                            >
                              Ver Detalhes
                            </button>
                          </div>
                        );
                      })}

                    {requests.filter(r => r.sector === selectedSector && r.isReturn && !r.deletedAt).length === 0 && (
                      <div className="p-12 text-center text-slate-500 space-y-2">
                        <RotateCcw className="mx-auto text-slate-300" size={40} />
                        <p className="font-bold text-sm text-slate-700">Nenhuma solicitação de devolução enviada.</p>
                        <p className="text-xs text-slate-500">Utilize o botão acima "Solicitar Devolução" para registrar o retorno de algum material ao almoxarifado.</p>
                      </div>
                    )}
                  </div>
                </div>
              )}

              {devolutionSubTab === 'sector_stock' && (
                <div className="bg-white rounded-3xl border border-slate-200/80 shadow-xs overflow-hidden">
                  <div className="p-6 border-b border-slate-100">
                    <h3 className="text-base font-black text-slate-900">Itens em Estoque no Setor ({selectedSector})</h3>
                    <p className="text-xs text-slate-500 font-medium">Selecione qualquer material guardado neste setor (inclusive vencidos) para devolver ao almoxarifado.</p>
                  </div>

                  <div className="p-4 sm:p-6 space-y-3">
                    {(() => {
                      const currentSectorStock = items.filter(i => 
                        !i.deletedAt && 
                        (i.location === selectedSector || (selectedSector === 'Farmácia' && i.location === 'Farmácia')) && 
                        i.quantity > 0
                      );

                      if (currentSectorStock.length === 0) {
                        return (
                          <div className="p-12 text-center text-slate-500 space-y-2">
                            <Package className="mx-auto text-slate-300" size={40} />
                            <p className="font-bold text-sm text-slate-700">Nenhum item em estoque no setor.</p>
                            <p className="text-xs text-slate-500">Quando a farmácia/setor possuir saldo em estoque, os itens aparecerão aqui para devolução imediata.</p>
                          </div>
                        );
                      }

                      return currentSectorStock.map(item => {
                        const expired = isExpired(item);
                        const nearExpiry = isNearExpiry(item);

                        return (
                          <div 
                            key={item.id} 
                            className={`border rounded-2xl p-4 sm:p-5 transition-all flex flex-col sm:flex-row sm:items-center justify-between gap-4 ${
                              expired 
                                ? 'bg-rose-50/50 border-rose-200 hover:border-rose-300' 
                                : nearExpiry 
                                ? 'bg-amber-50/40 border-amber-200 hover:border-amber-300' 
                                : 'bg-slate-50/60 border-slate-200/60 hover:border-amber-200'
                            }`}
                          >
                            <div className="space-y-1.5 flex-1">
                              <div className="flex flex-wrap items-center gap-2">
                                <span className="font-black text-sm text-slate-900">{item.name}</span>
                                {item.category && (
                                  <span className="text-[10px] font-bold text-slate-500 bg-white px-2 py-0.5 rounded-md border border-slate-200">
                                    {item.category}
                                  </span>
                                )}
                                {expired ? (
                                  <span className="text-[10px] font-black text-rose-700 bg-rose-100 border border-rose-300 px-2.5 py-0.5 rounded-full uppercase tracking-wider flex items-center gap-1">
                                    <AlertTriangle size={12} /> Vencido
                                  </span>
                                ) : nearExpiry ? (
                                  <span className="text-[10px] font-black text-amber-700 bg-amber-100 border border-amber-300 px-2.5 py-0.5 rounded-full uppercase tracking-wider flex items-center gap-1">
                                    <AlertTriangle size={12} /> Validade Próxima
                                  </span>
                                ) : (
                                  <span className="text-[10px] font-black text-emerald-700 bg-emerald-100 border border-emerald-300 px-2 py-0.5 rounded-full uppercase tracking-wider">
                                    Em dia
                                  </span>
                                )}
                              </div>

                              <p className="text-xs text-slate-500 font-medium">
                                Quantidade no Estoque: <span className="font-bold text-slate-900">{item.quantity} {item.unit_measure || 'unid'}</span>
                                {item.batch_number && (
                                  <> - Lote: <span className="font-bold text-slate-700">{item.batch_number}</span></>
                                )}
                                {item.expiry_date && item.expiry_date !== 'Indeterminada' && (
                                  <> - Validade: <span className={`font-bold ${expired ? 'text-rose-600' : 'text-slate-700'}`}>{new Date(item.expiry_date).toLocaleDateString('pt-BR')}</span></>
                                )}
                              </p>
                            </div>

                            <button 
                              onClick={() => {
                                setDevolutionBasket([{
                                  product_id: item.id,
                                  product_name: item.name,
                                  quantity: item.quantity,
                                  maxQty: item.quantity,
                                  selectedBatchId: item.id
                                }]);
                                setDevolutionReason(expired ? 'Vencido' : 'Não teve uso');
                                setDevolutionObservation(expired ? `Material vencido em ${new Date(item.expiry_date).toLocaleDateString('pt-BR')}` : '');
                                setShowDevolutionModal({ show: true });
                              }}
                              className={`px-4 py-2.5 rounded-xl text-xs font-black transition-all shadow-xs self-end sm:self-center whitespace-nowrap flex items-center gap-1.5 ${
                                expired 
                                  ? 'bg-rose-600 hover:bg-rose-700 text-white' 
                                  : 'bg-amber-600 hover:bg-amber-700 text-white'
                              }`}
                            >
                              <RotateCcw size={14} /> Devolver ao Almoxarifado
                            </button>
                          </div>
                        );
                      });
                    })()}
                  </div>
                </div>
              )}

              {devolutionSubTab === 'eligible_deliveries' && (
                <div className="bg-white rounded-3xl border border-slate-200/80 shadow-xs overflow-hidden">
                  <div className="p-6 border-b border-slate-100">
                    <h3 className="text-base font-black text-slate-900">Entregas Realizadas ao Setor</h3>
                    <p className="text-xs text-slate-500 font-medium">Selecione uma das entregas recebidas abaixo para selecionar itens e devolver.</p>
                  </div>

                  <div className="p-4 sm:p-6 space-y-3">
                    {requests
                      .filter(r => r.sector === selectedSector && r.status === 'ENTREGUE' && !r.deletedAt)
                      .map(req => {
                        const reqItems = allRequestItems.filter(ri => ri.request_id === req.id);
                        return (
                          <div 
                            key={req.id} 
                            className="bg-slate-50/60 border border-slate-200/60 hover:border-amber-200 rounded-2xl p-4 sm:p-5 transition-all flex flex-col sm:flex-row sm:items-center justify-between gap-4"
                          >
                            <div className="space-y-2 flex-1">
                              <div className="flex items-center gap-2">
                                <span className="font-black text-xs text-slate-900 bg-white px-2.5 py-1 rounded-lg border border-slate-200">
                                  Entrega #{req.id.slice(-5).toUpperCase()}
                                </span>
                                <span className="text-xs text-slate-500 font-semibold">
                                  - {req.deliveredAt 
                                    ? `Entregue em: ${new Date(req.deliveredAt).toLocaleDateString('pt-BR')}` 
                                    : `Criada em: ${new Date(req.date).toLocaleDateString('pt-BR')}`}
                                </span>
                              </div>

                              <div className="flex flex-wrap gap-1.5 pt-1">
                                {reqItems.map(item => {
                                  const alreadyReturned = item.quantity_returned || 0;
                                  const remaining = item.quantity_approved - alreadyReturned;
                                  return (
                                    <span 
                                      key={item.id} 
                                      className={`text-xs font-bold px-2.5 py-1 rounded-lg border ${
                                        remaining <= 0 
                                          ? 'bg-slate-100 text-slate-400 line-through border-slate-200' 
                                          : 'bg-white text-slate-800 border-amber-200/80 shadow-2xs'
                                      }`}
                                    >
                                      {item.product_name} ({remaining}/{item.quantity_approved} disp.)
                                    </span>
                                  );
                                })}
                              </div>
                            </div>

                            <button 
                              onClick={() => {
                                const basketItems = reqItems.map(ri => {
                                  const alreadyReturned = ri.quantity_returned || 0;
                                  const remaining = ri.quantity_approved - alreadyReturned;
                                  const productBatches = items.filter(item => !item.deletedAt && item.name === ri.product_name);
                                  return {
                                    product_id: ri.product_id,
                                    product_name: ri.product_name,
                                    quantity: remaining,
                                    maxQty: remaining,
                                    selectedBatchId: ri.batch_id || productBatches[0]?.id || ''
                                  };
                                }).filter(item => item.quantity > 0);

                                if (basketItems.length === 0) {
                                  if (showToast) showToast("Todos os itens desta entrega já foram totalmente devolvidos.", "info");
                                  return;
                                }

                                setDevolutionBasket(basketItems);
                                setDevolutionReason('Não teve uso');
                                setDevolutionObservation('');
                                setShowDevolutionModal({ show: true, request: req });
                              }}
                              className="bg-amber-600 text-white px-4 py-2.5 rounded-xl text-xs font-black hover:bg-amber-700 transition-all shadow-sm whitespace-nowrap self-end sm:self-center"
                            >
                              Devolver Materiais
                            </button>
                          </div>
                        );
                      })}

                    {requests.filter(r => r.sector === selectedSector && r.status === 'ENTREGUE' && !r.deletedAt).length === 0 && (
                      <div className="p-12 text-center text-slate-500 space-y-2">
                        <RotateCcw className="mx-auto text-slate-300" size={40} />
                        <p className="font-bold text-sm text-slate-700">Nenhuma entrega elegível encontrada.</p>
                        <p className="text-xs text-slate-500">Seu setor precisa ter entregas concluídas ("ENTREGUE") no sistema para devolvê-las ao estoque.</p>
                      </div>
                    )}
                  </div>
                </div>
              )}
    </motion.div>
  );
};
