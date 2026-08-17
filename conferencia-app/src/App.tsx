import { useEffect, useState, useMemo, useCallback } from 'react';
import { supabase } from './supabase';
import { format, parseISO } from 'date-fns';
import { ptBR } from 'date-fns/locale';
import { 
  RefreshCw, ListChecks, ArrowRightLeft, FileSpreadsheet, 
  Trash2, Search, ArrowRight, ShieldAlert, BadgeCheck, CheckCircle2,
  DollarSign, CreditCard, Banknote, Landmark, Sparkles, Calendar,
  Lock, LogOut
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { ResponsiveContainer, Tooltip, Legend, LineChart, Line, XAxis, YAxis, CartesianGrid } from 'recharts';

// --- Types ---
type PCLABRow = {
  id: number;
  descricao: string;
  tipo: string;
  forma_pgto: string;
  dt_lancamento: string;
  vr_lanc: number;
  created_at: string;
}

type AjusteRow = {
  id: number;
  data: string;
  descricao: string;
  valor: number;
  categoria: string;
  lado: string;
  criado_em: string;
}

const MONTHS = [
  { value: 0, label: 'Janeiro' },
  { value: 1, label: 'Fevereiro' },
  { value: 2, label: 'Março' },
  { value: 3, label: 'Abril' },
  { value: 4, label: 'Maio' },
  { value: 5, label: 'Junho' },
  { value: 6, label: 'Julho' },
  { value: 7, label: 'Agosto' },
  { value: 8, label: 'Setembro' },
  { value: 9, label: 'Outubro' },
  { value: 10, label: 'Novembro' },
  { value: 11, label: 'Dezembro' }
];

const YEARS = [2025, 2026, 2027];

type MaquininhaRow = {
  id: string;
  dt_transacao: string;
  vr_recebido: number;
  metodo_pgto: string;
  created_at: string;
}

type ReceitaBBRow = {
  id: number;
  data_lancamento: string;
  valor: number;
  descricao_historico: string;
  descricao_complementar: string;
  descricao_normalizada: string;
  numero_documento: string;
  classificacao: string | null;
}

type ClassificacaoRow = {
  descricao_normalizada: string;
  classificacao: string;
  criado_em: string;
}

type ReconciledTarget = {
  id: string | number;
  data: string;
  valor: number;
  descricao: string;
  tipo: string;
}

type MatchItem = {
  pclab: PCLABRow[];
  target: ReconciledTarget[];
  isGroup?: boolean;
  isCrossDate?: boolean;
  isForced?: boolean;
}

type UnclassifiedGroup = {
  descricao_normalizada: string;
  quantidade: number;
  data_ultimo: string;
  valor_ultimo: number;
  total_valor: number;
  items: ReceitaBBRow[];
}

// --- Utils ---
function formatCurrency(val: number) {
  return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(val);
}

function formatDate(dateStr: string) {
  if (!dateStr) return '';
  // Timezone-safe parse for YYYY-MM-DD
  const match = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) {
    return `${match[3]}/${match[2]}/${match[1]}`;
  }
  try {
    return format(parseISO(dateStr), 'dd/MM/yyyy', { locale: ptBR });
  } catch (e) {
    return dateStr;
  }
}

function formatDateHour(dateStr: string) {
  if (!dateStr) return '';
  // Timezone-safe parse for YYYY-MM-DD HH:MM
  const match = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})/);
  if (match) {
    return `${match[3]}/${match[2]}/${match[1]} ${match[4]}:${match[5]}`;
  }
  try {
    return format(parseISO(dateStr), 'dd/MM/yyyy HH:mm', { locale: ptBR });
  } catch (e) {
    return dateStr;
  }
}

// Helper to find subset sum match (PCLAB vs Maquininha/Extrato by day)
function findGroupMatches(pclabItems: PCLABRow[], targetItems: ReconciledTarget[]) {
  const matched: MatchItem[] = [];
  
  // Track matched status in local mutable lists
  const pList = pclabItems.map(p => ({ ...p, matched: false }));
  const tList = targetItems.map(t => ({ ...t, matched: false }));

  // Helper to find subsets of tList that sum to targetVal
  const findSubsetSumTarget = (items: typeof tList, targetVal: number): typeof tList | null => {
    const n = items.length;
    const limit = Math.min(n, 7);
    for (let i = 1; i < (1 << limit); i++) {
      const subset: typeof tList = [];
      let sum = 0;
      for (let j = 0; j < limit; j++) {
        if ((i & (1 << j)) > 0) {
          subset.push(items[j]);
          sum += items[j].valor;
        }
      }
      if (subset.length >= 2 && Math.abs(sum - targetVal) < 0.05) {
        return subset;
      }
    }
    return null;
  };

  // Helper to find subsets of pList that sum to targetVal
  const findSubsetSumPclab = (items: typeof pList, targetVal: number): typeof pList | null => {
    const n = items.length;
    const limit = Math.min(n, 7);
    for (let i = 1; i < (1 << limit); i++) {
      const subset: typeof pList = [];
      let sum = 0;
      for (let j = 0; j < limit; j++) {
        if ((i & (1 << j)) > 0) {
          subset.push(items[j]);
          sum += Number(items[j].vr_lanc);
        }
      }
      if (subset.length >= 2 && Math.abs(sum - targetVal) < 0.05) {
        return subset;
      }
    }
    return null;
  };

  // Step 1: Match 1 PCLAB item vs multiple Target items
  pList.forEach(p => {
    if (p.matched) return;
    const availableTargets = tList.filter(t => !t.matched);
    const subset = findSubsetSumTarget(availableTargets, Number(p.vr_lanc));
    if (subset) {
      p.matched = true;
      subset.forEach(s => {
        const found = tList.find(x => x.id === s.id);
        if (found) found.matched = true;
      });
      matched.push({
        pclab: [p],
        target: subset,
        isGroup: true
      });
    }
  });

  // Step 2: Match 1 Target item vs multiple PCLAB items
  tList.forEach(t => {
    if (t.matched) return;
    const availablePclab = pList.filter(p => !p.matched);
    const subset = findSubsetSumPclab(availablePclab, t.valor);
    if (subset) {
      t.matched = true;
      subset.forEach(s => {
        const found = pList.find(x => x.id === s.id);
        if (found) found.matched = true;
      });
      matched.push({
        pclab: subset,
        target: [t],
        isGroup: true
      });
    }
  });

  const unmatchedP = pList.filter(p => !p.matched);
  const unmatchedT = tList.filter(t => !t.matched);

  return { matched, unmatchedP, unmatchedT };
}

// --- Helpers for Faturamento classification ---
const isPersonalPedro = (desc: string, comp: string) => {
  const d = (desc || '').toLowerCase();
  const c = (comp || '').toLowerCase();
  return d.includes('03618401108') || c.includes('03618401108') ||
    d.includes('00003618401108') || c.includes('00003618401108') ||
    (d.includes('pedro henrique') && !d.includes('ltda')) ||
    (c.includes('pedro henrique') && !c.includes('ltda'));
};

const isDisregardedReceipt = (r: ReceitaBBRow) => {
  const desc = (r.descricao_normalizada || '').toLowerCase();
  const comp = (r.descricao_complementar || '').toLowerCase();

  // 1. Check if it matches Pedro's personal CPF/name - if so, it is NOT disregarded
  if (isPersonalPedro(r.descricao_normalizada, r.descricao_complementar)) {
    return false;
  }
  
  // 2. Pedro H C Dias LTDA
  if (desc.includes('57650247000170') || comp.includes('57650247000170') ||
      desc.includes('pedro h c dias') || comp.includes('pedro h c dias') ||
      desc.includes('pedro h. c. dias') || comp.includes('pedro h. c. dias') ||
      desc.includes('pedro h.c. dias') || comp.includes('pedro h.c. dias') ||
      desc.includes('pedro h c d') || comp.includes('pedro h c d')) {
    return true;
  }
  
  // 3. Laura de Souza Sisdelli
  if (desc.includes('38248524000179') || comp.includes('38248524000179') ||
      desc.includes('03743327112') || comp.includes('03743327112') ||
      desc.includes('laura de so') || comp.includes('laura de so') ||
      desc.includes('laura de souza') || comp.includes('laura de souza') ||
      desc.includes('laura sisdelli') || comp.includes('laura sisdelli') ||
      desc.includes('laura de souza sisdelli') || comp.includes('laura de souza sisdelli')) {
    return true;
  }
  
  // 4. Henrique Sisdelli Castro
  if (desc.includes('12038859116') || comp.includes('12038859116') ||
      desc.includes('henrique si') || comp.includes('henrique si') ||
      desc.includes('henrique sisde') || comp.includes('henrique sisde') ||
      desc.includes('henrique sisdelli') || comp.includes('henrique sisdelli') ||
      desc.includes('henrique sisdelli castro') || comp.includes('henrique sisdelli castro')) {
    return true;
  }
  
  return r.classificacao?.toLowerCase() === 'desconsiderar';
};

const getCategoryForReceipt = (r: ReceitaBBRow) => {
  const desc = (r.descricao_normalizada || '').toLowerCase();
  const comp = (r.descricao_complementar || '').toLowerCase();
  const hist = (r.descricao_historico || '').toLowerCase();
  const fullText = `${desc} ${comp} ${hist}`;

  if (fullText.includes('cargill') || fullText.includes('10249419000135')) {
    return 'Cargill';
  }

  if (fullText.includes('unimed') || fullText.includes('33546979000157')) {
    return 'Unimed';
  }

  if (fullText.includes('chromatox') || fullText.includes('14877243000117') || fullText.includes('laboratorio ch') || fullText.includes('chomatox') || fullText.includes('chroma')) {
    return 'Chromatox';
  }

  if (isDisregardedReceipt(r)) {
    return 'Desconsiderado';
  }

  return 'Particular';
};

export default function App() {
  // Auth State
  const [session, setSession] = useState<any>(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [email, setEmail] = useState('labsaopauloinaciolandia@gmail.com');
  const [password, setPassword] = useState('');
  const [loginError, setLoginError] = useState('');
  const [loginLoading, setLoginLoading] = useState(false);

  const [loading, setLoading] = useState(true);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'conferencia' | 'classificacao' | 'regras' | 'ajustes' | 'faturamento'>('dashboard');
  
  // Data State
  const [pclabRows, setPclabRows] = useState<PCLABRow[]>([]);
  const [maquininhaRows, setMaquininhaRows] = useState<MaquininhaRow[]>([]);
  const [receitasBbRows, setReceitasBbRows] = useState<ReceitaBBRow[]>([]);
  const [regrasClassificacao, setRegrasClassificacao] = useState<ClassificacaoRow[]>([]);
  const [ajustesRows, setAjustesRows] = useState<AjusteRow[]>([]);
  const [conferenciaLoading, setConferenciaLoading] = useState(false);

  // Sub-tab selection for conference tab
  const [confSubTab, setConfSubTab] = useState<'cartao' | 'pix'>('cartao');
  // Checkbox states for force reconciliation
  const [selectedPclabIds, setSelectedPclabIds] = useState<(number | string)[]>([]);
  const [selectedTargetIds, setSelectedTargetIds] = useState<(number | string)[]>([]);

  // Auth Listener
  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session);
      setAuthLoading(false);
    });

    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
    });

    return () => subscription.unsubscribe();
  }, []);

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoginError('');
    setLoginLoading(true);
    try {
      const { error } = await supabase.auth.signInWithPassword({ email, password });
      if (error) {
        setLoginError(error.message === 'Invalid login credentials' ? 'E-mail ou senha incorretos.' : error.message);
      }
    } catch (err: any) {
      setLoginError('Erro de conexão com o servidor.');
    } finally {
      setLoginLoading(false);
    }
  };

  const handleLogout = async () => {
    await supabase.auth.signOut();
  };

  useEffect(() => {
    setSelectedPclabIds([]);
    setSelectedTargetIds([]);
  }, [confSubTab, activeTab]);

  const handleTogglePclab = (id: number | string) => {
    setSelectedPclabIds(prev => 
      prev.includes(id) ? prev.filter(x => x !== id) : [...prev, id]
    );
  };

  const handleToggleTarget = (id: number | string) => {
    setSelectedTargetIds(prev => 
      prev.includes(id) ? prev.filter(x => x !== id) : [...prev, id]
    );
  };

  const handleForceReconcile = async () => {
    if (selectedPclabIds.length + selectedTargetIds.length < 1) {
      alert("Selecione pelo menos 1 transação para conciliar forçado.");
      return;
    }

    let dateStr = format(new Date(), 'yyyy-MM-dd');
    if (selectedPclabIds.length > 0) {
      const firstPclab = pclabRows.find(p => String(p.id) === String(selectedPclabIds[0]));
      if (firstPclab) dateStr = firstPclab.dt_lancamento.substring(0, 10);
    } else if (selectedTargetIds.length > 0) {
      const firstMaq = maquininhaRows.find(m => String(m.id) === String(selectedTargetIds[0]));
      if (firstMaq) {
        dateStr = firstMaq.dt_transacao.substring(0, 10);
      } else {
        const firstBB = receitasBbRows.find(r => String(r.id) === String(selectedTargetIds[0]));
        if (firstBB) dateStr = firstBB.data_lancamento.substring(0, 10);
      }
    }

    const descStr = `[Forçado] pclab:${selectedPclabIds.join(',')} | destino:${selectedTargetIds.join(',')}`;

    setConferenciaLoading(true);
    try {
      const { error } = await supabase
        .from('Conferencia_Ajustes')
        .insert({
          data: dateStr,
          descricao: descStr,
          valor: 0,
          categoria: confSubTab,
          lado: 'pclab'
        });

      if (error) throw error;

      setSelectedPclabIds([]);
      setSelectedTargetIds([]);
      await fetchData();
      alert("Conciliação forçada registrada com sucesso!");
    } catch (err: any) {
      console.error("Erro ao conciliar forçado:", err);
      alert("Erro ao realizar conciliação forçada: " + err.message);
    } finally {
      setConferenciaLoading(false);
    }
  };
  // Filters
  const [confFiltroStatus, setConfFiltroStatus] = useState<'todos' | 'conciliados' | 'divergencias'>('divergencias');
  const [confBusca, setConfBusca] = useState('');
  const [confFiltroData, setConfFiltroData] = useState<string>('');

  // Date Filters
  const [startDate, setStartDate] = useState<string>(() => {
    const d = new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    return `${y}-${m}-01`;
  });
  const [endDate, setEndDate] = useState<string>(() => {
    const d = new Date();
    const y = d.getFullYear();
    const m = d.getMonth() + 1;
    const lastDay = new Date(y, m, 0).getDate();
    const mStr = String(m).padStart(2, '0');
    return `${y}-${mStr}-${String(lastDay).padStart(2, '0')}`;
  });

  // Month & Year Selector (supports multiple selection)
  const [selectedMonths, setSelectedMonths] = useState<number[]>([new Date().getMonth()]);
  const [selectedYears, setSelectedYears] = useState<number[]>([new Date().getFullYear()]);
  const [useCustomPeriod, setUseCustomPeriod] = useState<boolean>(false);
  const [useCompletePeriod, setUseCompletePeriod] = useState<boolean>(false);

  // Cash Reconciliation specific state
  
  // Calendar specific state
  const [selectedCalendarDate, setSelectedCalendarDate] = useState<string | null>(null);

  // Form State for new adjustment
  const [ajustFormDate, setAjustFormDate] = useState<string>('');
  const [ajustFormDesc, setAjustFormDesc] = useState<string>('');
  const [ajustFormVal, setAjustFormVal] = useState<string>('');
  const [ajustFormCat, setAjustFormCat] = useState<string>('dinheiro');
  const [ajustFormLado, setAjustFormLado] = useState<string>('pclab');
  const [savingAjust, setSavingAjust] = useState<boolean>(false);

  // Classification selection
  const [savingClass, setSavingClass] = useState<string | null>(null);

  // Faturamento tab specific state
  const [faturamentoSubTab, setFaturamentoSubTab] = useState<'particular' | 'unimed' | 'cargill' | 'chromatox' | 'desconsiderados'>('particular');
  const [faturamentoBusca, setFaturamentoBusca] = useState('');

  useEffect(() => {
    fetchData();
  }, []);

  // When active tab is adjusted or startDate is set, update date in form to match
  useEffect(() => {
    if (startDate) setAjustFormDate(startDate);
  }, [startDate]);

  async function fetchData() {
    setConferenciaLoading(true);
    try {
      const { data: pclab } = await supabase.from('Conferencia_PCLAB').select('*');
      const { data: maquininha } = await supabase.from('Conferencia_Maquininha').select('*');
      const { data: receitasBB } = await supabase.from('Conferencia_ReceitasBB').select('*');
      const { data: regras } = await supabase.from('Conferencia_ReceitasClassificacao').select('*').order('criado_em', { ascending: false });
      const { data: ajustes } = await supabase.from('Conferencia_Ajustes').select('*').order('data', { ascending: false });

      if (pclab) {
        setPclabRows(pclab);
      }
      if (maquininha) setMaquininhaRows(maquininha);
      if (ajustes) setAjustesRows(ajustes);

      // Pré-processar receitasBB para desconsiderar automaticamente lançamentos do CNPJ da empresa ou sócios
      if (receitasBB) {
        const preprocessedBB = receitasBB.map(r => {
          const desc = (r.descricao_normalizada || '').toLowerCase();
          const comp = (r.descricao_complementar || '').toLowerCase();

          // 1. Pedro H C Dias LTDA (Corporate)
          const isCorporatePedro = 
            desc.includes('57650247000170') || comp.includes('57650247000170') ||
            desc.includes('pedro h c dias') || comp.includes('pedro h c dias') ||
            desc.includes('pedro h. c. dias') || comp.includes('pedro h. c. dias') ||
            desc.includes('pedro h.c. dias') || comp.includes('pedro h.c. dias') ||
            desc.includes('pedro h c d') || comp.includes('pedro h c d');

          // 2. Laura de Souza Sisdelli
          const isLaura = 
            desc.includes('38248524000179') || comp.includes('38248524000179') ||
            desc.includes('03743327112') || comp.includes('03743327112') ||
            desc.includes('laura de so') || comp.includes('laura de so') ||
            desc.includes('laura de souza') || comp.includes('laura de souza') ||
            desc.includes('laura sisdelli') || comp.includes('laura sisdelli');

          // 3. Henrique Sisdelli Castro
          const isHenrique = 
            desc.includes('12038859116') || comp.includes('12038859116') ||
            desc.includes('henrique si') || comp.includes('henrique si') ||
            desc.includes('henrique sisde') || comp.includes('henrique sisde') ||
            desc.includes('henrique sisdelli') || comp.includes('henrique sisdelli');

          if (isCorporatePedro || isLaura || isHenrique) {
            return { ...r, classificacao: 'desconsiderar' };
          }

          // Se for o Pedro pessoal, garanta que NÃO seja desconsiderado (mudar para Transferência Bancária)
          const isPersonal = isPersonalPedro(r.descricao_normalizada || '', r.descricao_complementar || '');

          if (isPersonal && (r.classificacao === 'desconsiderar' || !r.classificacao)) {
            return { ...r, classificacao: 'Transferencia Bancária' };
          }

          return r;
        });
        setReceitasBbRows(preprocessedBB);
      }

      // Ocultar regras que contenham Pedro para manter a lista limpa
      if (regras) {
        const filteredRegras = regras.filter(r => 
          !(r.descricao_normalizada || '').toLowerCase().includes('pedro')
        );
        setRegrasClassificacao(filteredRegras);
      }
    } catch (e) {
      console.error("Erro ao carregar dados:", e);
    } finally {
      setConferenciaLoading(false);
      setLoading(false);
    }
  }

  const isInPeriod = useCallback((dateStr: string) => {
    if (!dateStr) return false;
    
    if (useCompletePeriod) {
      return true;
    }
    
    if (useCustomPeriod) {
      const date = dateStr.substring(0, 10);
      if (startDate && date < startDate) return false;
      if (endDate && date > endDate) return false;
      return true;
    }

    // Filter by multiple selected months & years (timezone-safe parsing)
    const yearMatch = dateStr.match(/^(\d{4})-(\d{2})/);
    if (!yearMatch) return false;
    const year = parseInt(yearMatch[1], 10);
    const month = parseInt(yearMatch[2], 10) - 1;

    if (selectedMonths.length > 0 && !selectedMonths.includes(month)) return false;
    if (selectedYears.length > 0 && !selectedYears.includes(year)) return false;

    return true;
  }, [useCompletePeriod, useCustomPeriod, startDate, endDate, selectedMonths, selectedYears]);

  const reconciliationData = useMemo(() => {
    const getOnlyDate = (dateStr: string) => {
      if (!dateStr) return '';
      return dateStr.substring(0, 10);
    };


    const parseForcedMatch = (desc: string) => {
      const regex = /^\[Forçado\]\s*pclab:([^|]*)\|\s*destino:(.*)$/;
      const match = desc.match(regex);
      if (!match) return null;
      const pclabStr = match[1].trim();
      const destinoStr = match[2].trim();
      const pclabIds = pclabStr ? pclabStr.split(',').map(s => s.trim()).filter(Boolean) : [];
      const destinoIds = destinoStr ? destinoStr.split(',').map(s => s.trim()).filter(Boolean) : [];
      return { pclabIds, destinoIds };
    };

    const periodAjustes = ajustesRows.filter(a => isInPeriod(a.data) && (!a.descricao || !a.descricao.startsWith('[Forçado]')));

    // Parse forced matches globally
    const parsedForcedMatches = (ajustesRows || []).map(a => {
      if (!a.descricao || !a.descricao.startsWith('[Forçado]')) return null;
      const match = parseForcedMatch(a.descricao);
      if (!match) return null;
      return {
        id: a.id,
        categoria: a.categoria,
        pclabIds: match.pclabIds,
        destinoIds: match.destinoIds
      };
    }).filter(Boolean);

    // --- 1. Cartões ---
    const virtualPclabCards = periodAjustes.filter(a => 
      a.categoria === 'cartao' && a.lado === 'pclab'
    ).map(a => ({
      id: -Number(a.id),
      descricao: `[Ajuste Manual] ${a.descricao}`,
      tipo: 'Ajuste',
      forma_pgto: 'Cartão',
      dt_lancamento: `${a.data}T00:00:00`,
      vr_lanc: Number(a.valor),
      created_at: a.criado_em,
      matched: false
    }));

    const virtualMaqCards = periodAjustes.filter(a => 
      a.categoria === 'cartao' && a.lado === 'destino'
    ).map(a => ({
      id: `ajuste-${a.id}`,
      dt_transacao: `${a.data}T00:00:00`,
      vr_recebido: Number(a.valor),
      metodo_pgto: `[Ajuste Manual] ${a.descricao}`,
      created_at: a.criado_em,
      matched: false
    }));

    const pclabCards = [
      ...pclabRows.filter(r => 
        isInPeriod(r.dt_lancamento) && (
          r.forma_pgto?.toLowerCase().includes('cartã') || 
          r.forma_pgto?.toLowerCase().includes('cartao')
        )
      ).map(r => ({ ...r, matched: false })),
      ...virtualPclabCards
    ];

    const maqCards = [
      ...maquininhaRows.filter(r => 
        isInPeriod(r.dt_transacao) && (
          r.metodo_pgto?.toLowerCase().includes('cartã') || 
          r.metodo_pgto?.toLowerCase().includes('cartao')
        )
      ).map(r => ({ ...r, matched: false })),
      ...virtualMaqCards
    ];

    const cardsMatched: MatchItem[] = [];

    // Nível 0: Conciliação Forçada por Ajuste
    parsedForcedMatches.forEach(f => {
      if (f && f.categoria === 'cartao') {
        const matchedP = pclabCards.filter(p => f.pclabIds.includes(String(p.id)));
        const matchedT = maqCards.filter(m => f.destinoIds.includes(String(m.id)));
        
        if (matchedP.length > 0 || matchedT.length > 0) {
          matchedP.forEach(p => p.matched = true);
          matchedT.forEach(t => t.matched = true);
          cardsMatched.push({
            pclab: matchedP,
            target: matchedT.map(t => ({
              id: t.id,
              data: t.dt_transacao,
              valor: Number(t.vr_recebido),
              descricao: t.metodo_pgto,
              tipo: 'Maquininha'
            })),
            isForced: true
          });
        }
      }
    });

    // Nível 1.1: Pareamento Exato 1-para-1 (mesmo dia e valor)
    pclabCards.forEach(p => {
      if (p.matched) return;
      const pDate = getOnlyDate(p.dt_lancamento);
      const pVal = Number(p.vr_lanc);

      const match = maqCards.find(m => 
        !m.matched && 
        getOnlyDate(m.dt_transacao) === pDate && 
        Math.abs(Number(m.vr_recebido) - pVal) < 0.01
      );

      if (match) {
        p.matched = true;
        match.matched = true;
        cardsMatched.push({ 
          pclab: [p], 
          target: [{
            id: match.id,
            data: match.dt_transacao,
            valor: Number(match.vr_recebido),
            descricao: match.metodo_pgto,
            tipo: 'Maquininha'
          }]
        });
      }
    });

    // Nível 1.2: Ajuste de Conferência por Lote/Grupo de cartões no mesmo dia
    const remainingCardsPclab = pclabCards.filter(p => !p.matched);
    const remainingCardsTargetMapped: ReconciledTarget[] = maqCards.filter(m => !m.matched).map(m => ({
      id: m.id,
      data: m.dt_transacao,
      valor: Number(m.vr_recebido),
      descricao: m.metodo_pgto,
      tipo: 'Maquininha'
    }));

    // Agrupar por dia os restantes não pareados
    const pclabUnmatchedByDay = new Map<string, PCLABRow[]>();
    remainingCardsPclab.forEach(p => {
      const day = getOnlyDate(p.dt_lancamento);
      if (!pclabUnmatchedByDay.has(day)) pclabUnmatchedByDay.set(day, []);
      pclabUnmatchedByDay.get(day)!.push(p);
    });

    const targetUnmatchedByDay = new Map<string, ReconciledTarget[]>();
    remainingCardsTargetMapped.forEach(t => {
      const day = getOnlyDate(t.data);
      if (!targetUnmatchedByDay.has(day)) targetUnmatchedByDay.set(day, []);
      targetUnmatchedByDay.get(day)!.push(t);
    });

    const cardsFinalMatched = [...cardsMatched];
    const tempCardsUnmatchedPclab: (PCLABRow & { matched?: boolean })[] = [];
    const tempCardsUnmatchedTarget: ReconciledTarget[] = [];

    // Pegar todos os dias únicos
    const allCardDays = new Set<string>([
      ...Array.from(pclabUnmatchedByDay.keys()),
      ...Array.from(targetUnmatchedByDay.keys())
    ]);

    allCardDays.forEach(day => {
      const dayP = pclabUnmatchedByDay.get(day) || [];
      const dayT = targetUnmatchedByDay.get(day) || [];

      if (dayP.length > 0 && dayT.length > 0) {
        const { matched, unmatchedP, unmatchedT } = findGroupMatches(dayP, dayT);
        cardsFinalMatched.push(...matched);
        tempCardsUnmatchedPclab.push(...unmatchedP);
        tempCardsUnmatchedTarget.push(...unmatchedT);
      } else {
        tempCardsUnmatchedPclab.push(...dayP);
        tempCardsUnmatchedTarget.push(...dayT);
      }
    });

    // Nível 1.3: Pareamento Cruzado de Datas desativado (conciliação apenas no mesmo dia)
    const cardsFinalUnmatchedPclab = tempCardsUnmatchedPclab;
    const cardsFinalUnmatchedTarget = tempCardsUnmatchedTarget;

    // --- 2. Pix & Transferências ---
    const virtualPclabPix = periodAjustes.filter(a => 
      a.categoria === 'pix' && a.lado === 'pclab'
    ).map(a => ({
      id: -Number(a.id),
      descricao: `[Ajuste Manual] ${a.descricao}`,
      tipo: 'Ajuste',
      forma_pgto: 'Transf. Bancária',
      dt_lancamento: `${a.data}T00:00:00`,
      vr_lanc: Number(a.valor),
      created_at: a.criado_em,
      matched: false
    }));

    const virtualMaqPix = periodAjustes.filter(a => 
      a.categoria === 'pix' && a.lado === 'destino'
    ).map(a => ({
      id: `ajuste-${a.id}`,
      dt_transacao: `${a.data}T00:00:00`,
      vr_recebido: Number(a.valor),
      metodo_pgto: `[Ajuste Manual] ${a.descricao}`,
      created_at: a.criado_em,
      matched: false
    }));

    const pclabTransf = [
      ...pclabRows.filter(r => 
        isInPeriod(r.dt_lancamento) && r.forma_pgto === 'Transf. Bancária'
      ).map(r => ({ ...r, matched: false })),
      ...virtualPclabPix
    ];

    const maqPix = [
      ...maquininhaRows.filter(r => 
        isInPeriod(r.dt_transacao) && r.metodo_pgto === 'Pix'
      ).map(r => ({ ...r, matched: false })),
      ...virtualMaqPix
    ];

    const bbTransf = receitasBbRows.filter(r => 
      isInPeriod(r.data_lancamento) && (
        r.classificacao?.toLowerCase() === 'transferencia bancária' ||
        r.classificacao?.toLowerCase() === 'transferencia bancaria'
      )
    ).map(r => ({ ...r, matched: false }));

    const pixMatched: MatchItem[] = [];

    // Nível 0: Conciliação Forçada por Ajuste (Pix)
    parsedForcedMatches.forEach(f => {
      if (f && f.categoria === 'pix') {
        const matchedP = pclabTransf.filter(p => f.pclabIds.includes(String(p.id)));
        const matchedMaq = maqPix.filter(m => f.destinoIds.includes(String(m.id)));
        const matchedBB = bbTransf.filter(b => f.destinoIds.includes(String(b.id)));
        
        if (matchedP.length > 0 || matchedMaq.length > 0 || matchedBB.length > 0) {
          matchedP.forEach(p => p.matched = true);
          matchedMaq.forEach(m => m.matched = true);
          matchedBB.forEach(b => b.matched = true);
          
          const targetItems: ReconciledTarget[] = [
            ...matchedMaq.map(m => ({
              id: m.id,
              data: m.dt_transacao,
              valor: Number(m.vr_recebido),
              descricao: m.metodo_pgto,
              tipo: 'Maquininha (Pix)'
            })),
            ...matchedBB.map(b => ({
              id: b.id,
              data: b.data_lancamento,
              valor: Number(b.valor),
              descricao: `${b.descricao_historico} - ${b.descricao_normalizada}`,
              tipo: 'Banco (Receita)'
            }))
          ];
          
          pixMatched.push({
            pclab: matchedP,
            target: targetItems,
            isForced: true
          });
        }
      }
    });

    // Auto-conciliação automática de recebimentos bancários de Pedro
    bbTransf.forEach(b => {
      if (!b.matched && isPersonalPedro(b.descricao_normalizada || '', b.descricao_complementar || '')) {
        b.matched = true;
        pixMatched.push({
          pclab: [],
          target: [{
            id: b.id,
            data: b.data_lancamento,
            valor: Number(b.valor),
            descricao: `[Auto-Conciliado Pedro] ${b.descricao_historico} - ${b.descricao_normalizada}`,
            tipo: 'Banco (Receita)'
          }]
        });
      }
    });

    // Passo A: PCLAB Transf. Bancária vs Maquininha Pix (mesmo dia e valor)
    pclabTransf.forEach(p => {
      if (p.matched) return;
      const pDate = getOnlyDate(p.dt_lancamento);
      const pVal = Number(p.vr_lanc);

      const match = maqPix.find(m => 
        !m.matched && 
        getOnlyDate(m.dt_transacao) === pDate && 
        Math.abs(Number(m.vr_recebido) - pVal) < 0.01
      );

      if (match) {
        p.matched = true;
        match.matched = true;
        pixMatched.push({ 
          pclab: [p], 
          target: [{
            id: match.id,
            data: match.dt_transacao,
            valor: Number(match.vr_recebido),
            descricao: match.metodo_pgto,
            tipo: 'Maquininha (Pix)'
          }]
        });
      }
    });

    // Passo B: Restante PCLAB Transf. Bancária vs ReceitasBB (mesmo dia e valor)
    pclabTransf.filter(p => !p.matched).forEach(p => {
      const pDate = getOnlyDate(p.dt_lancamento);
      const pVal = Number(p.vr_lanc);

      const match = bbTransf.find(b => 
        !b.matched && 
        getOnlyDate(b.data_lancamento) === pDate && 
        Math.abs(Number(b.valor) - pVal) < 0.01
      );

      if (match) {
        p.matched = true;
        match.matched = true;
        pixMatched.push({ 
          pclab: [p], 
          target: [{
            id: match.id,
            data: match.data_lancamento,
            valor: Number(match.valor),
            descricao: `${match.descricao_historico} - ${match.descricao_normalizada}`,
            tipo: 'Banco (Receita)'
          }]
        });
      }
    });

    const tempPixUnmatchedPclab = pclabTransf.filter(p => !p.matched);
    const tempPixUnmatchedTarget: ReconciledTarget[] = [
      ...maqPix.filter(m => !m.matched).map(m => ({
        id: m.id,
        data: m.dt_transacao,
        valor: Number(m.vr_recebido),
        descricao: m.metodo_pgto,
        tipo: 'Maquininha (Pix)'
      })),
      ...bbTransf.filter(b => !b.matched).map(b => ({
        id: b.id,
        data: b.data_lancamento,
        valor: Number(b.valor),
        descricao: `${b.descricao_historico} - ${b.descricao_normalizada}`,
        tipo: 'Banco (Receita)'
      }))
    ];

    // Passo C: Pareamento Cruzado de Datas desativado (conciliação apenas no mesmo dia)
    const pixUnmatchedPclab = tempPixUnmatchedPclab;
    const pixUnmatchedTarget = tempPixUnmatchedTarget;

    // --- 3. Dinheiro (Confrontado Globalmente por Saldo / Sem Pareamento Individual) ---
    const virtualPclabDinheiro = periodAjustes.filter(a => 
      a.categoria === 'dinheiro' && a.lado === 'pclab'
    ).map(a => ({
      id: -Number(a.id),
      descricao: `[Ajuste Manual] ${a.descricao}`,
      tipo: 'Ajuste',
      forma_pgto: 'Dinheiro',
      dt_lancamento: `${a.data}T00:00:00`,
      vr_lanc: Number(a.valor),
      created_at: a.criado_em
    }));

    const virtualBbDinheiro = periodAjustes.filter(a => 
      a.categoria === 'dinheiro' && a.lado === 'destino'
    ).map(a => ({
      id: `ajuste-${a.id}`,
      data: a.data,
      valor: Number(a.valor),
      descricao: `[Ajuste Manual] ${a.descricao}`,
      tipo: 'Banco (Ajuste)'
    }));

    const pclabDinheiro = [
      ...pclabRows.filter(r => 
        isInPeriod(r.dt_lancamento) && r.forma_pgto === 'Dinheiro'
      ),
      ...virtualPclabDinheiro
    ];

    const bbDinheiroMapped: ReconciledTarget[] = [
      ...receitasBbRows.filter(r => 
        isInPeriod(r.data_lancamento) && r.classificacao?.toLowerCase() === 'dinheiro'
      ).map(b => ({
        id: b.id,
        data: b.data_lancamento,
        valor: Number(b.valor),
        descricao: b.descricao_historico || 'Depósito ATM',
        tipo: 'Banco (Dinheiro)'
      })),
      ...virtualBbDinheiro
    ];

    return {
      cartao: {
        matched: cardsFinalMatched,
        unmatchedPclab: cardsFinalUnmatchedPclab,
        unmatchedTarget: cardsFinalUnmatchedTarget
      },
      pix: {
        matched: pixMatched,
        unmatchedPclab: pixUnmatchedPclab,
        unmatchedTarget: pixUnmatchedTarget
      },
      dinheiro: {
        matched: [],
        unmatchedPclab: pclabDinheiro,
        unmatchedTarget: bbDinheiroMapped
      }
    };
  }, [pclabRows, maquininhaRows, receitasBbRows, startDate, endDate, ajustesRows, selectedMonths, selectedYears, useCustomPeriod, useCompletePeriod, isInPeriod]);

  // Filtered lists for the active sub-tab (for cards and pix)
  const currentSubTabData = useMemo(() => {
    const data = reconciliationData[confSubTab];
    
    // Helper to check if string matches search query
    const matchesSearch = (str: string, val: number) => {
      if (!confBusca) return true;
      const query = confBusca.toLowerCase();
      return (
        str.toLowerCase().includes(query) || 
        val.toString().includes(query) ||
        formatCurrency(val).includes(query)
      );
    };

    // Helper to extract date YYYY-MM-DD
    const getOnlyDate = (dateStr: string) => {
      if (!dateStr) return '';
      return dateStr.substring(0, 10);
    };

    // Helper to check if item's date matches selected date filter
    const matchesDate = (dateStr: string) => {
      if (!confFiltroData) return true;
      return getOnlyDate(dateStr) === confFiltroData;
    };

    const filteredMatched = data.matched.filter(m => {
      const pDate = m.pclab.length > 0 ? m.pclab[0].dt_lancamento : '';
      const tDate = m.target.length > 0 ? m.target[0].data : '';
      const itemDate = pDate || tDate || '';
      return matchesDate(itemDate) && (
        m.pclab.some(p => matchesSearch(p.descricao, p.vr_lanc)) ||
        m.target.some(t => matchesSearch(t.descricao, t.valor))
      );
    });

    const filteredUnmatchedPclab = data.unmatchedPclab.filter(p => 
      matchesDate(p.dt_lancamento) && matchesSearch(p.descricao, p.vr_lanc)
    );

    const filteredUnmatchedTarget = data.unmatchedTarget.filter(t => 
      matchesDate(t.data) && matchesSearch(t.descricao, t.valor)
    );

    return {
      matched: filteredMatched,
      unmatchedPclab: filteredUnmatchedPclab,
      unmatchedTarget: filteredUnmatchedTarget
    };
  }, [reconciliationData, confSubTab, confBusca, confFiltroData]);

  const sortedReconciliationRows = useMemo(() => {
    const getMatchDate = (item: MatchItem) => {
      const pDate = item.pclab.length > 0 ? item.pclab[0].dt_lancamento : '';
      const tDate = item.target.length > 0 ? item.target[0].data : '';
      return pDate || tDate || '';
    };

    const rows: (
      | { type: 'matched'; data: MatchItem; date: string }
      | { type: 'pclab'; data: PCLABRow; date: string }
      | { type: 'target'; data: ReconciledTarget; date: string }
    )[] = [];

    const showMatched = confFiltroStatus === 'todos' || confFiltroStatus === 'conciliados';
    const showDivergencias = confFiltroStatus === 'todos' || confFiltroStatus === 'divergencias';

    if (showMatched) {
      currentSubTabData.matched.forEach(m => {
        rows.push({ type: 'matched', data: m, date: getMatchDate(m) });
      });
    }

    if (showDivergencias) {
      currentSubTabData.unmatchedPclab.forEach(p => {
        rows.push({ type: 'pclab', data: p, date: p.dt_lancamento });
      });
      currentSubTabData.unmatchedTarget.forEach(t => {
        rows.push({ type: 'target', data: t, date: t.data });
      });
    }

    // Sort by date descending (newest first)
    return rows.sort((a, b) => b.date.localeCompare(a.date));
  }, [currentSubTabData, confFiltroStatus]);


  // Find the last update time in the database (PCLAB and Maquininha, ignoring checking account)
  const lastUpdateTime = useMemo(() => {
    let maxTime: Date | null = null;

    const checkRow = (row: any) => {
      const timeStr = row.imported_at || row.created_at;
      if (timeStr) {
        const d = new Date(timeStr);
        if (!isNaN(d.getTime())) {
          if (!maxTime || d > maxTime) {
            maxTime = d;
          }
        }
      }
    };

    pclabRows.forEach(checkRow);
    maquininhaRows.forEach(checkRow);

    return maxTime;
  }, [pclabRows, maquininhaRows]);

  const formatLastUpdate = (date: Date | null) => {
    if (!date) return '';
    const pad = (n: number) => String(n).padStart(2, '0');
    const day = pad(date.getDate());
    const month = pad(date.getMonth() + 1);
    const year = date.getFullYear();
    const hours = pad(date.getHours());
    const minutes = pad(date.getMinutes());
    return `${day}/${month}/${year} ${hours}:${minutes}`;
  };

  // Find minSelectedDate to calculate prior balance
  const minSelectedDate = useMemo(() => {
    if (useCustomPeriod) return startDate;
    if (useCompletePeriod) return '';
    if (selectedMonths.length === 0 || selectedYears.length === 0) return '';
    
    const sortedYears = [...selectedYears].sort((a, b) => a - b);
    const minYear = sortedYears[0];
    
    const sortedMonths = [...selectedMonths].sort((a, b) => a - b);
    const minMonth = sortedMonths[0];
    
    const pad = (n: number) => n.toString().padStart(2, '0');
    return `${minYear}-${pad(minMonth + 1)}-01`;
  }, [selectedMonths, selectedYears, useCustomPeriod, useCompletePeriod, startDate]);

  const startBalance = useMemo(() => {
    if (useCompletePeriod) return 0;
    
    let limitDate = '';
    if (useCustomPeriod) {
      limitDate = startDate || '';
    } else if (selectedMonths.length > 0 && selectedYears.length > 0) {
      const minYear = Math.min(...selectedYears);
      const minMonth = Math.min(...selectedMonths);
      const pad = (n: number) => n.toString().padStart(2, '0');
      limitDate = `${minYear}-${pad(minMonth + 1)}-01`;
    }
    
    if (!limitDate) return 0;
    
    // Sum faturamentos before limitDate
    let fatSum = 0;
    pclabRows.forEach(p => {
      if (p.forma_pgto === 'Dinheiro') {
        const date = p.dt_lancamento.substring(0, 10);
        if (date < limitDate) {
          fatSum += Number(p.vr_lanc);
        }
      }
    });
    ajustesRows.forEach(a => {
      if (a.categoria === 'dinheiro' && a.lado === 'pclab') {
        if (a.data < limitDate) {
          fatSum += Number(a.valor);
        }
      }
    });
    
    // Sum deposits before limitDate
    let depSum = 0;
    receitasBbRows.forEach(b => {
      if (b.classificacao?.toLowerCase() === 'dinheiro') {
        const date = b.data_lancamento.substring(0, 10);
        if (date < limitDate) {
          depSum += Number(b.valor);
        }
      }
    });
    ajustesRows.forEach(a => {
      if (a.categoria === 'dinheiro' && a.lado === 'destino') {
        if (a.data < limitDate) {
          depSum += Number(a.valor);
        }
      }
    });
    
    return depSum - fatSum;
  }, [useCompletePeriod, useCustomPeriod, startDate, selectedMonths, selectedYears, pclabRows, receitasBbRows, ajustesRows]);

  // Dinheiro Balance Calculations
  const dinheiroTotalPclab = useMemo(() => {
    return reconciliationData.dinheiro.unmatchedPclab.reduce((acc, curr) => acc + Number(curr.vr_lanc), 0);
  }, [reconciliationData.dinheiro]);

  const dinheiroTotalTarget = useMemo(() => {
    return reconciliationData.dinheiro.unmatchedTarget.reduce((acc, curr) => acc + Number(curr.valor), 0);
  }, [reconciliationData.dinheiro]);



  // Cash Ledger (Chronological unified feed for cash)
  const cashLedger = useMemo(() => {
    const faturamentos = reconciliationData.dinheiro.unmatchedPclab.map(p => ({
      id: p.id,
      date: p.dt_lancamento,
      type: 'saida' as const,
      description: p.descricao,
      value: Number(p.vr_lanc),
      source: 'PCLAB'
    }));

    const depositos = reconciliationData.dinheiro.unmatchedTarget.map(t => ({
      id: t.id,
      date: t.data,
      type: 'entrada' as const,
      description: t.descricao,
      value: Number(t.valor),
      source: 'Banco'
    }));

    const merged = [...faturamentos, ...depositos];

    // Sort chronologically. If same day, put 'entrada' first.
    merged.sort((a, b) => {
      const dateA = a.date.substring(0, 10);
      const dateB = b.date.substring(0, 10);
      if (dateA !== dateB) {
        return dateA.localeCompare(dateB);
      }
      if (a.type !== b.type) {
        return a.type === 'entrada' ? -1 : 1;
      }
      return a.date.localeCompare(b.date);
    });

    // Calculate running balance starting from startBalance
    let running = startBalance;
    return merged.map(item => {
      if (item.type === 'entrada') {
        running += item.value;
      } else {
        running -= item.value;
      }
      return {
        ...item,
        runningBalance: running
      };
    });
  }, [reconciliationData.dinheiro, startBalance]);



  // Daily aggregations for Cash (Dinheiro)
  const faturamentoByDay = useMemo(() => {
    const map = new Map<string, number>();
    pclabRows.forEach(p => {
      if (p.forma_pgto === 'Dinheiro') {
        const date = p.dt_lancamento.substring(0, 10);
        map.set(date, (map.get(date) || 0) + Number(p.vr_lanc));
      }
    });
    // Virtual adjustments on PCLAB side
    ajustesRows.forEach(a => {
      if (a.categoria === 'dinheiro' && a.lado === 'pclab') {
        map.set(a.data, (map.get(a.data) || 0) + Number(a.valor));
      }
    });
    return map;
  }, [pclabRows, ajustesRows]);

  const depositsByDay = useMemo(() => {
    const map = new Map<string, number>();
    receitasBbRows.forEach(b => {
      if (b.classificacao?.toLowerCase() === 'dinheiro') {
        const date = b.data_lancamento.substring(0, 10);
        map.set(date, (map.get(date) || 0) + Number(b.valor));
      }
    });
    // Virtual adjustments on BB side
    ajustesRows.forEach(a => {
      if (a.categoria === 'dinheiro' && a.lado === 'destino') {
        map.set(a.data, (map.get(a.data) || 0) + Number(a.valor));
      }
    });
    return map;
  }, [receitasBbRows, ajustesRows]);

  const calendarMonthsToRender = useMemo(() => {
    if (useCompletePeriod) {
      const ymSet = new Set<string>();
      pclabRows.forEach(r => {
        if (r.forma_pgto === 'Dinheiro') {
          ymSet.add(r.dt_lancamento.substring(0, 7));
        }
      });
      receitasBbRows.forEach(r => {
        if (r.classificacao?.toLowerCase() === 'dinheiro') {
          ymSet.add(r.data_lancamento.substring(0, 7));
        }
      });
      const sorted = Array.from(ymSet).sort();
      return sorted.map(ym => {
        const [y, m] = ym.split('-');
        return { year: parseInt(y, 10), month: parseInt(m, 10) - 1 };
      });
    }

    if (useCustomPeriod) {
      if (!startDate || !endDate) return [];
      const start = parseISO(startDate);
      const end = parseISO(endDate);
      const list = [];
      let curr = new Date(start.getFullYear(), start.getMonth(), 1);
      while (curr <= end) {
        list.push({ year: curr.getFullYear(), month: curr.getMonth() });
        curr.setMonth(curr.getMonth() + 1);
      }
      return list;
    }

    const list: { year: number; month: number }[] = [];
    const sortedYears = [...selectedYears].sort((a, b) => a - b);
    const sortedMonths = [...selectedMonths].sort((a, b) => a - b);
    
    sortedYears.forEach(y => {
      sortedMonths.forEach(m => {
        list.push({ year: y, month: m });
      });
    });
    return list;
  }, [selectedMonths, selectedYears, useCustomPeriod, useCompletePeriod, startDate, endDate, pclabRows, receitasBbRows]);

  const getCalendarDayStats = (dateStr: string) => {
    const faturamento = faturamentoByDay.get(dateStr) || 0;
    const deposits = depositsByDay.get(dateStr) || 0;

    let lastTx = null;
    for (let i = cashLedger.length - 1; i >= 0; i--) {
      if (cashLedger[i].date.substring(0, 10) <= dateStr) {
        lastTx = cashLedger[i];
        break;
      }
    }
    
    const runningBalance = lastTx ? lastTx.runningBalance : startBalance;
    const hasTransactions = faturamento > 0 || deposits > 0;
    const isReconciled = runningBalance >= 0;

    return {
      faturamento,
      deposits,
      runningBalance,
      isReconciled,
      hasTransactions
    };
  };
  const selectedDayTransactions = useMemo(() => {
    if (!selectedCalendarDate) return [];
    
    const faturamentos = pclabRows.filter(p => 
      p.forma_pgto === 'Dinheiro' && p.dt_lancamento.substring(0, 10) === selectedCalendarDate
    ).map(p => ({
      id: p.id,
      description: p.descricao,
      value: Number(p.vr_lanc),
      type: 'saida' as const,
      source: 'PCLAB'
    }));

    const adjustmentsPclab = ajustesRows.filter(a =>
      a.categoria === 'dinheiro' && a.lado === 'pclab' && a.data === selectedCalendarDate
    ).map(a => ({
      id: `ajuste-${a.id}`,
      description: `[Ajuste] ${a.descricao}`,
      value: Number(a.valor),
      type: 'saida' as const,
      source: 'Ajuste PCLAB'
    }));

    const depositos = receitasBbRows.filter(b =>
      b.classificacao?.toLowerCase() === 'dinheiro' && b.data_lancamento.substring(0, 10) === selectedCalendarDate
    ).map(b => ({
      id: b.id,
      description: b.descricao_historico || 'Depósito',
      value: Number(b.valor),
      type: 'entrada' as const,
      source: 'Banco'
    }));

    const adjustmentsBB = ajustesRows.filter(a =>
      a.categoria === 'dinheiro' && a.lado === 'destino' && a.data === selectedCalendarDate
    ).map(a => ({
      id: `ajuste-${a.id}`,
      description: `[Ajuste] ${a.descricao}`,
      value: Number(a.valor),
      type: 'entrada' as const,
      source: 'Ajuste Banco'
    }));

    return [...faturamentos, ...adjustmentsPclab, ...depositos, ...adjustmentsBB];
  }, [pclabRows, receitasBbRows, ajustesRows, selectedCalendarDate]);

  function renderCalendar(year: number, month: number) {
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const firstDayIndex = new Date(year, month, 1).getDay();
    const daysOfWeek = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];
    const cells = [];
    
    for (let i = 0; i < firstDayIndex; i++) {
      cells.push(<div key={`empty-${i}`} className="calendar-day empty"></div>);
    }
    
    for (let day = 1; day <= daysInMonth; day++) {
      const pad = (n: number) => n.toString().padStart(2, '0');
      const dateStr = `${year}-${pad(month + 1)}-${pad(day)}`;
      const stats = getCalendarDayStats(dateStr);
      const isSelected = selectedCalendarDate === dateStr;
      
      let dayClass = "calendar-day";
      if (stats.hasTransactions) {
        dayClass += stats.isReconciled ? " reconciled" : " discrepancy";
      } else {
        dayClass += " no-tx";
      }
      if (isSelected) {
        dayClass += " selected";
      }
      
      cells.push(
        <div 
          key={day} 
          className={dayClass}
          onClick={() => setSelectedCalendarDate(dateStr)}
          style={{ cursor: 'pointer' }}
        >
          <div className="day-number">{day}</div>
          {stats.hasTransactions ? (
            <div className="day-stats">
              {stats.deposits > 0 && (
                <div className="stat-in">+{formatCurrency(stats.deposits)}</div>
              )}
              {stats.faturamento > 0 && (
                <div className="stat-out">-{formatCurrency(stats.faturamento)}</div>
              )}
              <div className="stat-bal" style={{ whiteSpace: 'nowrap' }}>S: {formatCurrency(stats.runningBalance)}</div>
            </div>
          ) : (
            <div className="day-stats no-data">
              <div className="stat-bal-neutral">{formatCurrency(stats.runningBalance)}</div>
            </div>
          )}
        </div>
      );
    }

    const totalCells = cells.length;
    const remaining = totalCells % 7;
    if (remaining > 0) {
      for (let i = 0; i < 7 - remaining; i++) {
        cells.push(<div key={`empty-end-${i}`} className="calendar-day empty"></div>);
      }
    }

    const monthName = MONTHS.find(m => m.value === month)?.label || '';

    return (
      <div key={`${year}-${month}`} className="calendar-container glass-card" style={{ padding: '1rem', marginBottom: '1.5rem', background: 'rgba(255,255,255,0.01)' }}>
        <h3 style={{ marginBottom: '1rem', textAlign: 'center', color: 'var(--text-dark)', fontWeight: 600 }}>
          {monthName} de {year}
        </h3>
        <div className="calendar-grid-header">
          {daysOfWeek.map(d => (
            <div key={d} className="calendar-grid-header-cell">{d}</div>
          ))}
        </div>
        <div className="calendar-grid">
          {cells}
        </div>
      </div>
    );
  }

  // Helper to describe the selected period
  const getPeriodDescription = () => {
    if (useCompletePeriod) return 'Período Completo';
    if (useCustomPeriod) return `De ${formatDate(startDate)} até ${formatDate(endDate)}`;
    
    const monthLabels = selectedMonths.length > 0 
      ? selectedMonths.map(m => MONTHS.find(x => x.value === m)?.label).join(', ') 
      : 'Todos os meses';
    const yearLabels = selectedYears.length > 0 
      ? selectedYears.join(', ') 
      : 'Todos os anos';
    return `${monthLabels} de ${yearLabels}`;
  };

  // Unclassified Receipts Group
  const unclassifiedGroups = useMemo<UnclassifiedGroup[]>(() => {
    const unclass = receitasBbRows.filter(r => !r.classificacao);
    
    const groups = new Map<string, ReceitaBBRow[]>();
    unclass.forEach(r => {
      const desc = r.descricao_normalizada || 'Sem descrição';
      if (!groups.has(desc)) groups.set(desc, []);
      groups.get(desc)!.push(r);
    });

    return Array.from(groups.entries()).map(([desc, items]) => {
      const sorted = [...items].sort((a, b) => b.data_lancamento.localeCompare(a.data_lancamento));
      return {
        descricao_normalizada: desc,
        quantidade: items.length,
        data_ultimo: sorted[0].data_lancamento,
        valor_ultimo: Number(sorted[0].valor),
        total_valor: items.reduce((acc, curr) => acc + Number(curr.valor), 0),
        items: items
      };
    }).sort((a, b) => b.quantidade - a.quantidade);
  }, [receitasBbRows]);

  // Handlers
  const handleClassificar = async (descricao: string, classificacao: string) => {
    setSavingClass(descricao);
    try {
      const { error } = await supabase
        .from('Conferencia_ReceitasClassificacao')
        .insert({ descricao_normalizada: descricao, classificacao });
      
      if (error) throw error;
      await fetchData();
    } catch (e) {
      console.error("Erro ao salvar classificação:", e);
      alert("Erro ao salvar classificação!");
    } finally {
      setSavingClass(null);
    }
  };

  const handleDeleteRegra = async (descricao: string) => {
    if (!confirm(`Deseja mesmo excluir a regra para: "${descricao}"?`)) return;
    try {
      const { error } = await supabase
        .from('Conferencia_ReceitasClassificacao')
        .delete()
        .eq('descricao_normalizada', descricao);

      if (error) throw error;
      await fetchData();
    } catch (e) {
      console.error("Erro ao deletar regra:", e);
      alert("Erro ao deletar regra!");
    }
  };

  // Toggle Month selection
  const toggleMonth = (monthVal: number) => {
    if (selectedMonths.includes(monthVal)) {
      if (selectedMonths.length > 1) {
        setSelectedMonths(selectedMonths.filter(m => m !== monthVal));
      }
    } else {
      setSelectedMonths([...selectedMonths, monthVal]);
    }
  };

  // Toggle Year selection
  const toggleYear = (yearVal: number) => {
    if (selectedYears.includes(yearVal)) {
      if (selectedYears.length > 1) {
        setSelectedYears(selectedYears.filter(y => y !== yearVal));
      }
    } else {
      setSelectedYears([...selectedYears, yearVal]);
    }
  };

  // Manual Adjustments Handlers
  const handleCreateAjuste = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!ajustFormDate || !ajustFormDesc || !ajustFormVal) {
      alert("Por favor, preencha todos os campos do ajuste.");
      return;
    }

    setSavingAjust(true);
    try {
      const { error } = await supabase
        .from('Conferencia_Ajustes')
        .insert({
          data: ajustFormDate,
          descricao: ajustFormDesc,
          valor: Number(ajustFormVal),
          categoria: ajustFormCat,
          lado: ajustFormLado
        });

      if (error) throw error;
      
      setAjustFormDesc('');
      setAjustFormVal('');
      await fetchData();
      alert("Ajuste de conciliação registrado com sucesso!");
    } catch (e) {
      console.error("Erro ao salvar ajuste:", e);
      alert("Erro ao registrar ajuste!");
    } finally {
      setSavingAjust(false);
    }
  };

  const handleDeleteAjuste = async (id: number) => {
    if (!confirm("Deseja realmente excluir este ajuste manual?")) return;
    try {
      const { error } = await supabase
        .from('Conferencia_Ajustes')
        .delete()
        .eq('id', id);

      if (error) throw error;
      await fetchData();
    } catch (e) {
      console.error("Erro ao deletar ajuste:", e);
      alert("Erro ao excluir ajuste!");
    }
  };

  const filteredAjustes = useMemo(() => {
    return ajustesRows.filter(a => {
      if (startDate && a.data < startDate) return false;
      if (endDate && a.data > endDate) return false;
      return true;
    });
  }, [ajustesRows, startDate, endDate]);

  const handleExportConferencia = (type: 'cartao' | 'pix' | 'dinheiro') => {
    if (type === 'dinheiro') {
      const data = reconciliationData.dinheiro;
      
      // 1. Livro Caixa
      const ledgerList = cashLedger.map(item => ({
        'Data': item.source === 'PCLAB' ? formatDateHour(item.date) : formatDate(item.date),
        'Origem': item.source,
        'Descrição': item.description,
        'Entrada (+)': item.type === 'entrada' ? item.value : 0,
        'Saída (-)': item.type === 'saida' ? item.value : 0,
        'Saldo Acumulado': item.runningBalance
      }));

      ledgerList.unshift({
        'Data': minSelectedDate ? formatDate(minSelectedDate) : 'Início',
        'Origem': 'SALDO INICIAL',
        'Descrição': 'Saldo Inicial de Caixa',
        'Entrada (+)': 0,
        'Saída (-)': 0,
        'Saldo Acumulado': startBalance
      });

      const pclabList = data.unmatchedPclab.map(p => ({
        'Data Lançamento': formatDateHour(p.dt_lancamento),
        'Descrição': p.descricao,
        'Valor Registrado': Number(p.vr_lanc)
      }));

      const bbList = data.unmatchedTarget.map(t => ({
        'Data Depósito': formatDate(t.data),
        'Histórico Banco': t.descricao,
        'Valor Depositado': Number(t.valor)
      }));

      const summary = [{
        'Saldo Inicial': startBalance,
        'Total PCLAB (Faturamento)': dinheiroTotalPclab,
        'Total Banco (Depósitos)': dinheiroTotalTarget,
        'Saldo Final': startBalance + dinheiroTotalTarget - dinheiroTotalPclab,
        'Status': (startBalance + dinheiroTotalTarget - dinheiroTotalPclab) >= 0 ? 'SALDO REGULAR' : 'SALDO NEGATIVO'
      }];

      const workbook = XLSX.utils.book_new();
      
      const sheetLedger = XLSX.utils.json_to_sheet(ledgerList);
      XLSX.utils.book_append_sheet(workbook, sheetLedger, "Livro Caixa (Extrato)");

      const sheetSummary = XLSX.utils.json_to_sheet(summary);
      XLSX.utils.book_append_sheet(workbook, sheetSummary, "Resumo Saldo Dinheiro");

      const sheetPclab = XLSX.utils.json_to_sheet(pclabList);
      XLSX.utils.book_append_sheet(workbook, sheetPclab, "Faturamento PCLAB (Dinheiro)");

      const sheetBB = XLSX.utils.json_to_sheet(bbList);
      XLSX.utils.book_append_sheet(workbook, sheetBB, "Depósitos Banco (Dinheiro)");

      const periodDesc = getPeriodDescription().replace(/, /g, '_').replace(/ /g, '_');
      XLSX.writeFile(workbook, `Conferencia_Dinheiro_Saldo_${periodDesc}_${format(new Date(), 'dd-MM-yyyy')}.xlsx`);
      return;
    }

    const subData = reconciliationData[type];
    const exportList: any[] = [];

    // Matched
    subData.matched.forEach((m: MatchItem, groupIndex: number) => {
      const maxLen = Math.max(m.pclab.length, m.target.length);
      for (let i = 0; i < maxLen; i++) {
        const p = m.pclab[i] || null;
        const t = m.target[i] || null;
        exportList.push({
          'Status': m.isCrossDate ? 'CONCILIADO (DATA DIF.)' : m.isGroup ? `CONCILIADO LOTE #${groupIndex + 1}` : 'CONCILIADO',
          'PCLAB Data': p ? formatDate(p.dt_lancamento) : '-',
          'PCLAB Descrição': p ? p.descricao : '-',
          'PCLAB Valor': p ? Number(p.vr_lanc) : 0,
          'Destino Tipo': t ? t.tipo : '-',
          'Destino Data': t ? formatDate(t.data) : '-',
          'Destino Descrição': t ? t.descricao : '-',
          'Destino Valor': t ? Number(t.valor) : 0,
          'Diferença': (i === 0 && !m.isGroup) ? 0 : '-'
        });
      }
    });

    // Unmatched PCLAB
    subData.unmatchedPclab.forEach((p: PCLABRow) => {
      exportList.push({
        'Status': 'PENDENTE PCLAB',
        'PCLAB Data': formatDate(p.dt_lancamento),
        'PCLAB Descrição': p.descricao,
        'PCLAB Valor': Number(p.vr_lanc),
        'Destino Tipo': '-',
        'Destino Data': '-',
        'Destino Descrição': 'Não encontrado no destino',
        'Destino Valor': 0,
        'Diferença': Number(p.vr_lanc)
      });
    });

    // Unmatched Target
    subData.unmatchedTarget.forEach((t: ReconciledTarget) => {
      exportList.push({
        'Status': 'PENDENTE DESTINO',
        'PCLAB Data': '-',
        'PCLAB Descrição': 'Não registrado no PCLAB',
        'PCLAB Valor': 0,
        'Destino Tipo': t.tipo,
        'Destino Data': formatDate(t.data),
        'Destino Descrição': t.descricao,
        'Destino Valor': Number(t.valor),
        'Diferença': Number(t.valor)
      });
    });

    const worksheet = XLSX.utils.json_to_sheet(exportList);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Conferência");
    XLSX.writeFile(workbook, `Conferencia_${type}_${format(new Date(), 'dd-MM-yyyy')}.xlsx`);
  };

  // --- Faturamento Memoized Data (Month-over-Month, BB Only) ---
  const faturamentoMonths = useMemo(() => {
    const monthsSet = new Set<string>();
    
    if (useCompletePeriod) {
      receitasBbRows.forEach(r => {
        if (r.data_lancamento) {
          monthsSet.add(r.data_lancamento.substring(0, 7));
        }
      });
    } else if (useCustomPeriod) {
      if (startDate && endDate) {
        let curr = new Date(parseISO(startDate));
        const end = new Date(parseISO(endDate));
        while (curr <= end) {
          const yyyy = curr.getFullYear();
          const mm = (curr.getMonth() + 1).toString().padStart(2, '0');
          monthsSet.add(`${yyyy}-${mm}`);
          curr.setMonth(curr.getMonth() + 1);
        }
      }
    } else {
      const sortedYears = [...selectedYears].sort((a, b) => a - b);
      const sortedMonths = [...selectedMonths].sort((a, b) => a - b);
      sortedYears.forEach(y => {
        sortedMonths.forEach(m => {
          const mm = (m + 1).toString().padStart(2, '0');
          monthsSet.add(`${y}-${mm}`);
        });
      });
    }

    return Array.from(monthsSet).sort();
  }, [useCompletePeriod, useCustomPeriod, startDate, endDate, selectedMonths, selectedYears, receitasBbRows]);

  const faturamentoData = useMemo(() => {
    const grid: { [month: string]: { cargill: number; unimed: number; chromatox: number; particular: number; disregarded: number } } = {};
    
    let totalCargill = 0;
    let totalUnimed = 0;
    let totalChromatox = 0;
    let totalParticular = 0;
    let totalDisregarded = 0;

    const cargillRows: ReceitaBBRow[] = [];
    const unimedRows: ReceitaBBRow[] = [];
    const chromatoxRows: ReceitaBBRow[] = [];
    const particularRows: ReceitaBBRow[] = [];
    const disregardedRows: ReceitaBBRow[] = [];

    faturamentoMonths.forEach(m => {
      grid[m] = { cargill: 0, unimed: 0, chromatox: 0, particular: 0, disregarded: 0 };
    });

    receitasBbRows.forEach(r => {
      const monthStr = r.data_lancamento?.substring(0, 7);
      if (!monthStr || !faturamentoMonths.includes(monthStr)) return;

      const category = getCategoryForReceipt(r);
      const val = Number(r.valor) || 0;

      if (category === 'Cargill') {
        grid[monthStr].cargill += val;
        totalCargill += val;
        cargillRows.push(r);
      } else if (category === 'Unimed') {
        grid[monthStr].unimed += val;
        totalUnimed += val;
        unimedRows.push(r);
      } else if (category === 'Chromatox') {
        grid[monthStr].chromatox += val;
        totalChromatox += val;
        chromatoxRows.push(r);
      } else if (category === 'Particular') {
        grid[monthStr].particular += val;
        totalParticular += val;
        particularRows.push(r);
      } else if (category === 'Desconsiderado') {
        grid[monthStr].disregarded += val;
        totalDisregarded += val;
        disregardedRows.push(r);
      }
    });

    return {
      grid,
      totals: {
        cargill: totalCargill,
        unimed: totalUnimed,
        chromatox: totalChromatox,
        particular: totalParticular,
        disregarded: totalDisregarded
      },
      rows: {
        cargill: cargillRows,
        unimed: unimedRows,
        chromatox: chromatoxRows,
        particular: particularRows,
        disregarded: disregardedRows
      }
    };
  }, [receitasBbRows, faturamentoMonths]);

  const formatMonthYear = (monthStr: string) => {
    const [year, month] = monthStr.split('-');
    const date = new Date(Number(year), Number(month) - 1, 1);
    return format(date, 'MMM/yyyy', { locale: ptBR });
  };

  const chartData = useMemo(() => {
    return faturamentoMonths.map(m => {
      const monthData = faturamentoData.grid[m] || { particular: 0, unimed: 0, cargill: 0, chromatox: 0 };
      return {
        name: formatMonthYear(m),
        Particular: monthData.particular,
        Unimed: monthData.unimed,
        Cargill: monthData.cargill,
        Chromatox: monthData.chromatox
      };
    });
  }, [faturamentoMonths, faturamentoData]);

  const filteredFaturamentoReceipts = useMemo(() => {
    let list: ReceitaBBRow[] = [];
    if (faturamentoSubTab === 'particular') {
      list = faturamentoData.rows.particular;
    } else if (faturamentoSubTab === 'unimed') {
      list = faturamentoData.rows.unimed;
    } else if (faturamentoSubTab === 'cargill') {
      list = faturamentoData.rows.cargill;
    } else if (faturamentoSubTab === 'chromatox') {
      list = faturamentoData.rows.chromatox;
    } else if (faturamentoSubTab === 'desconsiderados') {
      list = faturamentoData.rows.disregarded;
    }

    const query = faturamentoBusca.toLowerCase();
    const result = !faturamentoBusca 
      ? [...list] 
      : list.filter(r => {
          const desc = (r.descricao_normalizada || '').toLowerCase();
          const comp = (r.descricao_complementar || '').toLowerCase();
          const doc = (r.numero_documento || '').toLowerCase();
          const valStr = r.valor.toString();
          const valFormatted = formatCurrency(r.valor);
          
          return desc.includes(query) || comp.includes(query) || doc.includes(query) || valStr.includes(query) || valFormatted.includes(query);
        });

    return result.sort((a, b) => b.data_lancamento.localeCompare(a.data_lancamento));
  }, [faturamentoData, faturamentoSubTab, faturamentoBusca]);

  const renderLastColumnCell = (category: 'particular' | 'unimed' | 'cargill' | 'chromatox' | 'total') => {
    if (faturamentoMonths.length === 2) {
      const m1 = faturamentoMonths[0];
      const m2 = faturamentoMonths[1];
      
      let v1 = 0;
      let v2 = 0;
      
      if (category === 'total') {
        v1 = (faturamentoData.grid[m1]?.particular || 0) + (faturamentoData.grid[m1]?.unimed || 0) + (faturamentoData.grid[m1]?.cargill || 0) + (faturamentoData.grid[m1]?.chromatox || 0);
        v2 = (faturamentoData.grid[m2]?.particular || 0) + (faturamentoData.grid[m2]?.unimed || 0) + (faturamentoData.grid[m2]?.cargill || 0) + (faturamentoData.grid[m2]?.chromatox || 0);
      } else {
        v1 = faturamentoData.grid[m1]?.[category] || 0;
        v2 = faturamentoData.grid[m2]?.[category] || 0;
      }
      
      if (v1 === 0) {
        if (v2 === 0) return '0.00%';
        return <span className="val-positive" style={{ fontWeight: 600 }}>+100.00%</span>;
      }
      
      const diff = ((v2 - v1) / v1) * 100;
      const sign = diff > 0 ? '+' : '';
      const colorClass = diff > 0 ? 'val-positive' : diff < 0 ? 'val-negative' : '';
      
      return (
        <span className={colorClass} style={{ fontWeight: 600 }}>
          {sign}{diff.toFixed(2)}%
        </span>
      );
    }
    
    // Otherwise render total period sum
    let total = 0;
    if (category === 'total') {
      total = faturamentoData.totals.particular + faturamentoData.totals.unimed + faturamentoData.totals.cargill + faturamentoData.totals.chromatox;
    } else {
      total = faturamentoData.totals[category] || 0;
    }
    return formatCurrency(total);
  };

  const handleExportFaturamento = () => {
    let list: ReceitaBBRow[] = [];
    let name = '';
    
    if (faturamentoSubTab === 'particular') {
      list = faturamentoData.rows.particular;
      name = 'Particular';
    } else if (faturamentoSubTab === 'unimed') {
      list = faturamentoData.rows.unimed;
      name = 'Unimed';
    } else if (faturamentoSubTab === 'cargill') {
      list = faturamentoData.rows.cargill;
      name = 'Cargill';
    } else if (faturamentoSubTab === 'chromatox') {
      list = faturamentoData.rows.chromatox;
      name = 'Chromatox';
    } else if (faturamentoSubTab === 'desconsiderados') {
      list = faturamentoData.rows.disregarded;
      name = 'Desconsiderados';
    }

    const exportList = list.map(r => ({
      'Data Lançamento': formatDate(r.data_lancamento),
      'Histórico': r.descricao_historico,
      'Descrição Complementar': r.descricao_complementar,
      'Descrição Normalizada': r.descricao_normalizada,
      'Número Documento': r.numero_documento,
      'Valor': Number(r.valor)
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportList);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, `Faturamento ${name}`);
    XLSX.writeFile(workbook, `Faturamento_${name}_${format(new Date(), 'dd-MM-yyyy')}.xlsx`);
  };

  if (authLoading) {
    return (
      <div className="loader-container">
        <div className="spinner"></div>
      </div>
    );
  }

  if (!session) {
    return (
      <div className="login-wrapper">
        <div className="login-glass-card fade-in">
          <div className="login-header">
            <div className="logo-icon-container">
              <Lock size={32} />
            </div>
            <h2>Área Restrita</h2>
            <p>Conferência de Recebimentos - Laboratório São Paulo</p>
          </div>
          
          {loginError && (
            <div className="login-error-badge">
              <ShieldAlert size={16} />
              <span>{loginError}</span>
            </div>
          )}
          
          <form onSubmit={handleLogin} className="login-form">
            <div className="form-group">
              <label className="form-label">E-mail</label>
              <input 
                type="email" 
                className="input-text" 
                value={email}
                onChange={e => setEmail(e.target.value)}
                required
                placeholder="seu-email@gmail.com"
              />
            </div>
            
            <div className="form-group">
              <label className="form-label">Senha</label>
              <input 
                type="password" 
                className="input-text" 
                value={password}
                onChange={e => setPassword(e.target.value)}
                required
                placeholder="Digite sua senha"
              />
            </div>
            
            <button type="submit" className="btn-login" disabled={loginLoading}>
              {loginLoading ? <div className="spinner-small"></div> : 'Acessar Painel'}
            </button>
          </form>
        </div>
      </div>
    );
  }

  if (loading) {
    return (
      <div className="loader-container">
        <div className="spinner"></div>
      </div>
    );
  }

  return (
    <div className="app-container">
      {/* Header */}
      <header className="header">
        <div className="title-container">
          <h1>Conferência de Recebimentos</h1>
          <p>Confronto de dados: PCLAB vs Maquininha Ton & Extrato Banco do Brasil</p>
        </div>
        <div style={{ display: 'flex', gap: '1rem', alignItems: 'center', flexWrap: 'wrap' }}>
          {lastUpdateTime && (
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', fontSize: '0.75rem', color: 'var(--text-muted)' }}>
              <span>Última Atualização no Banco:</span>
              <strong style={{ color: 'var(--text-dark)' }}>{formatLastUpdate(lastUpdateTime)}</strong>
            </div>
          )}
          <button onClick={fetchData} className="btn-filter" style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', border: '1px solid var(--border-color)' }}>
            <RefreshCw size={16} className={conferenciaLoading ? 'spin-animation' : ''} /> Atualizar Dados
          </button>
          <button onClick={handleLogout} className="btn-logout">
            <LogOut size={16} /> Sair
          </button>
        </div>
      </header>

      {/* Filtro de Período Geral */}
      <div className="glass-card fade-in" style={{ marginBottom: '1.5rem', padding: '1.25rem', background: 'rgba(255, 255, 255, 0.02)', border: '1px solid var(--border-color)' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '1rem' }}>
          
          {/* Lado Esquerdo: Filtros de Mês e Ano */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem', flex: 1, minWidth: '300px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
              <Calendar size={16} style={{ color: 'var(--primary)' }} />
              <span style={{ fontWeight: 600, fontSize: '0.9rem', color: 'var(--text-dark)' }}>
                Filtro de Período:
              </span>
              {useCompletePeriod ? (
                <span className="badge badge-green" style={{ fontSize: '0.75rem' }}>Período Completo (Sem Filtros)</span>
              ) : useCustomPeriod ? (
                <span className="badge badge-blue" style={{ fontSize: '0.75rem' }}>Período Personalizado</span>
              ) : (
                <span className="badge badge-yellow" style={{ fontSize: '0.75rem' }}>Filtro por Meses/Anos</span>
              )}
            </div>

            {/* Renderização de Acordo com o Modo de Período */}
            {useCompletePeriod ? (
              <div style={{ color: 'var(--text-muted)', fontSize: '0.85rem', fontStyle: 'italic' }}>
                Exibindo todos os lançamentos carregados no sistema (sem filtros de data).
              </div>
            ) : useCustomPeriod ? (
              <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', flexWrap: 'wrap' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '0.25rem' }}>
                  <label style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>De:</label>
                  <input 
                    type="date" 
                    className="input-text" 
                    value={startDate} 
                    onChange={e => setStartDate(e.target.value)} 
                    style={{ width: 'auto', padding: '0.4rem 0.75rem', background: 'var(--bg-card)', border: '1px solid var(--border-color)', color: 'var(--text-dark)', borderRadius: '6px' }} 
                  />
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: '0.25rem' }}>
                  <label style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>Até:</label>
                  <input 
                    type="date" 
                    className="input-text" 
                    value={endDate} 
                    onChange={e => setEndDate(e.target.value)} 
                    style={{ width: 'auto', padding: '0.4rem 0.75rem', background: 'var(--bg-card)', border: '1px solid var(--border-color)', color: 'var(--text-dark)', borderRadius: '6px' }} 
                  />
                </div>
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
                {/* Pills de Meses */}
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.35rem', alignItems: 'center' }}>
                  <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)', width: '50px' }}>Meses:</span>
                  {MONTHS.map(m => {
                    const isSel = selectedMonths.includes(m.value);
                    return (
                      <button
                        key={m.value}
                        type="button"
                        onClick={() => toggleMonth(m.value)}
                        className="btn-filter"
                        style={{
                          padding: '0.25rem 0.6rem',
                          fontSize: '0.75rem',
                          borderRadius: '12px',
                          background: isSel ? 'var(--primary)' : 'rgba(255,255,255,0.03)',
                          color: isSel ? '#fff' : 'var(--text-muted)',
                          border: '1px solid ' + (isSel ? 'var(--primary)' : 'var(--border-color)'),
                          fontWeight: isSel ? 600 : 400
                        }}
                      >
                        {m.label.substring(0, 3)}
                      </button>
                    );
                  })}
                </div>

                {/* Pills de Anos */}
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.35rem', alignItems: 'center' }}>
                  <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)', width: '50px' }}>Anos:</span>
                  {YEARS.map(y => {
                    const isSel = selectedYears.includes(y);
                    return (
                      <button
                        key={y}
                        type="button"
                        onClick={() => toggleYear(y)}
                        className="btn-filter"
                        style={{
                          padding: '0.25rem 0.6rem',
                          fontSize: '0.75rem',
                          borderRadius: '12px',
                          background: isSel ? 'var(--primary)' : 'rgba(255,255,255,0.03)',
                          color: isSel ? '#fff' : 'var(--text-muted)',
                          border: '1px solid ' + (isSel ? 'var(--primary)' : 'var(--border-color)'),
                          fontWeight: isSel ? 600 : 400
                        }}
                      >
                        {y}
                      </button>
                    );
                  })}
                </div>
              </div>
            )}
          </div>

          {/* Lado Direito: Controles de Configuração de Período */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem', alignItems: 'flex-end' }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', cursor: 'pointer', fontSize: '0.85rem', color: 'var(--text-dark)', fontWeight: 500 }}>
              <input 
                type="checkbox" 
                checked={useCompletePeriod} 
                onChange={e => {
                  setUseCompletePeriod(e.target.checked);
                  if (e.target.checked) setUseCustomPeriod(false);
                }} 
                style={{ width: '16px', height: '16px', accentColor: 'var(--success)', cursor: 'pointer' }}
              />
              Período Completo
            </label>
            
            <label style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', cursor: 'pointer', fontSize: '0.85rem', color: 'var(--text-muted)' }}>
              <input 
                type="checkbox" 
                checked={useCustomPeriod} 
                onChange={e => {
                  setUseCustomPeriod(e.target.checked);
                  if (e.target.checked) setUseCompletePeriod(false);
                }} 
                style={{ width: '16px', height: '16px', accentColor: 'var(--primary)', cursor: 'pointer' }}
              />
              Período Personalizado
            </label>

            {!useCompletePeriod && !useCustomPeriod && (
              <button 
                type="button"
                className="btn-filter"
                style={{ padding: '0.35rem 0.6rem', fontSize: '0.75rem', marginTop: '0.25rem' }}
                onClick={() => {
                  const curM = new Date().getMonth();
                  const curY = new Date().getFullYear();
                  setSelectedMonths([curM]);
                  setSelectedYears([curY]);
                  const y = curY;
                  const m = String(curM + 1).padStart(2, '0');
                  const lastDay = new Date(y, curM + 1, 0).getDate();
                  setStartDate(`${y}-${m}-01`);
                  setEndDate(`${y}-${m}-${String(lastDay).padStart(2, '0')}`);
                }}
              >
                Resetar para Mês Atual
              </button>
            )}
          </div>

        </div>
      </div>

      {/* Nav Tabs */}
      <nav className="nav-tabs">
        <button className={`tab-btn ${activeTab === 'dashboard' ? 'active' : ''}`} onClick={() => setActiveTab('dashboard')}>
          <Banknote size={18} className="mr-2" /> Dinheiro
        </button>
        <button className={`tab-btn ${activeTab === 'conferencia' ? 'active' : ''}`} onClick={() => setActiveTab('conferencia')}>
          <ListChecks size={18} className="mr-2" /> Cartão / Pix
        </button>
        <button className={`tab-btn ${activeTab === 'faturamento' ? 'active' : ''}`} onClick={() => setActiveTab('faturamento')}>
          <FileSpreadsheet size={18} className="mr-2" /> Comparação Faturamento
        </button>
        <button className={`tab-btn ${activeTab === 'classificacao' ? 'active' : ''}`} onClick={() => setActiveTab('classificacao')}>
          <ArrowRightLeft size={18} className="mr-2" /> 
          Classificar Receitas ({unclassifiedGroups.length})
        </button>
        <button className={`tab-btn ${activeTab === 'regras' ? 'active' : ''}`} onClick={() => setActiveTab('regras')}>
          <Sparkles size={18} className="mr-2" /> Regras de Classificação ({regrasClassificacao.length})
        </button>
        <button className={`tab-btn ${activeTab === 'ajustes' ? 'active' : ''}`} onClick={() => setActiveTab('ajustes')}>
          <ArrowRightLeft size={18} className="mr-2" style={{ transform: 'rotate(90deg)' }} /> Ajustes Manuais ({ajustesRows.length})
        </button>
      </nav>

      {/* Tabs Content */}
      {activeTab === 'faturamento' && (
        <div className="fade-in">
          <div className="grid-12" style={{ gap: '1.5rem' }}>
            
            {/* Tabela de Comparação de Faturamento */}
            <div className="col-span-7-md" style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
              <div className="glass-card" style={{ padding: '1.5rem' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.25rem' }}>
                  <h2 style={{ color: 'var(--text-dark)', fontWeight: 600, fontSize: '1.25rem' }}>
                    Evolução Mensal de Faturamento (Banco do Brasil)
                  </h2>
                  <span className="badge badge-blue">Conta Corrente</span>
                </div>
                
                <div className="table-responsive">
                  <table className="table-conferencia">
                    <thead>
                      <tr>
                        <th>Categoria / Convênio</th>
                        {faturamentoMonths.map(m => (
                          <th key={m} style={{ textAlign: 'right' }}>{formatMonthYear(m)}</th>
                        ))}
                        <th style={{ textAlign: 'right', fontWeight: 700 }}>
                          {faturamentoMonths.length === 2 ? 'Diferença %' : 'Total Período'}
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {/* Particular */}
                      <tr>
                        <td style={{ fontWeight: 600 }}>Particular</td>
                        {faturamentoMonths.map(m => (
                          <td key={m} className="amount" style={{ textAlign: 'right' }}>
                            {formatCurrency(faturamentoData.grid[m].particular)}
                          </td>
                        ))}
                        <td className="amount" style={{ textAlign: 'right', fontWeight: 600 }}>
                          {renderLastColumnCell('particular')}
                        </td>
                      </tr>

                      {/* Unimed */}
                      <tr>
                        <td style={{ fontWeight: 600 }}>Unimed</td>
                        {faturamentoMonths.map(m => (
                          <td key={m} className="amount" style={{ textAlign: 'right' }}>
                            {formatCurrency(faturamentoData.grid[m].unimed)}
                          </td>
                        ))}
                        <td className="amount" style={{ textAlign: 'right', fontWeight: 600 }}>
                          {renderLastColumnCell('unimed')}
                        </td>
                      </tr>

                      {/* Cargill */}
                      <tr>
                        <td style={{ fontWeight: 600 }}>Cargill</td>
                        {faturamentoMonths.map(m => (
                          <td key={m} className="amount" style={{ textAlign: 'right' }}>
                            {formatCurrency(faturamentoData.grid[m].cargill)}
                          </td>
                        ))}
                        <td className="amount" style={{ textAlign: 'right', fontWeight: 600 }}>
                          {renderLastColumnCell('cargill')}
                        </td>
                      </tr>

                      {/* Chromatox */}
                      <tr>
                        <td style={{ fontWeight: 600 }}>Chromatox</td>
                        {faturamentoMonths.map(m => (
                          <td key={m} className="amount" style={{ textAlign: 'right' }}>
                            {formatCurrency(faturamentoData.grid[m].chromatox)}
                          </td>
                        ))}
                        <td className="amount" style={{ textAlign: 'right', fontWeight: 600 }}>
                          {renderLastColumnCell('chromatox')}
                        </td>
                      </tr>

                      {/* Total */}
                      <tr style={{ borderTop: '2px solid var(--border-color)', background: 'rgba(255,255,255,0.02)' }}>
                        <td style={{ fontWeight: 700, fontSize: '0.95rem' }}>Total Geral</td>
                        {faturamentoMonths.map(m => {
                          const monthSum = 
                            faturamentoData.grid[m].particular + 
                            faturamentoData.grid[m].unimed + 
                            faturamentoData.grid[m].cargill + 
                            faturamentoData.grid[m].chromatox;
                          return (
                            <td key={m} className="amount" style={{ textAlign: 'right', fontWeight: 700 }}>
                              {formatCurrency(monthSum)}
                            </td>
                          );
                        })}
                        <td className="amount" style={{ textAlign: 'right', fontWeight: 700, fontSize: '0.95rem' }}>
                          {renderLastColumnCell('total')}
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Gráfico Recharts */}
              <div className="glass-card" style={{ padding: '1.5rem' }}>
                <h3 style={{ color: 'var(--text-dark)', fontWeight: 600, fontSize: '1.1rem', marginBottom: '1rem' }}>
                  Evolução Mensal do Faturamento
                </h3>
                <div style={{ width: '100%', height: 260 }}>
                  <ResponsiveContainer>
                    <LineChart
                      data={chartData}
                      margin={{ top: 10, right: 20, left: 10, bottom: 5 }}
                    >
                      <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" />
                      <XAxis 
                        dataKey="name" 
                        stroke="var(--text-muted)" 
                        tick={{ fill: 'var(--text-muted)', fontSize: '0.8rem' }}
                      />
                      <YAxis 
                        stroke="var(--text-muted)" 
                        tick={{ fill: 'var(--text-muted)', fontSize: '0.8rem' }}
                        tickFormatter={(value) => `R$ ${value >= 1000 ? (value / 1000) + 'k' : value}`}
                      />
                      <Tooltip 
                        formatter={(value: any) => formatCurrency(Number(value))}
                        contentStyle={{
                          background: 'rgba(15, 23, 42, 0.9)',
                          border: '1px solid var(--border-color)',
                          borderRadius: '8px',
                          color: '#fff'
                        }}
                      />
                      <Legend 
                        verticalAlign="bottom" 
                        height={36} 
                        iconType="circle"
                        formatter={(value) => <span style={{ color: 'var(--text-dark)', fontSize: '0.85rem' }}>{value}</span>}
                      />
                      <Line type="monotone" dataKey="Particular" stroke="#10b981" strokeWidth={3} activeDot={{ r: 6 }} />
                      <Line type="monotone" dataKey="Unimed" stroke="#3b82f6" strokeWidth={3} activeDot={{ r: 6 }} />
                      <Line type="monotone" dataKey="Cargill" stroke="#f59e0b" strokeWidth={3} activeDot={{ r: 6 }} />
                      <Line type="monotone" dataKey="Chromatox" stroke="#8b5cf6" strokeWidth={3} activeDot={{ r: 6 }} />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>

            {/* Painel de Auditoria e Detalhes da Auditoria */}
            <div className="col-span-5-md" style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
              <div className="glass-card" style={{ padding: '1.5rem', display: 'flex', flexDirection: 'column', height: '100%' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem', flexWrap: 'wrap', gap: '0.5rem' }}>
                  <h3 style={{ color: 'var(--text-dark)', fontWeight: 600, fontSize: '1.1rem' }}>
                    Auditoria de Lançamentos
                  </h3>
                  <button onClick={handleExportFaturamento} className="btn-filter" style={{ display: 'flex', alignItems: 'center', gap: '0.25rem', padding: '0.35rem 0.6rem', fontSize: '0.75rem' }}>
                    <FileSpreadsheet size={14} /> Exportar Excel
                  </button>
                </div>

                {/* Sub-abas de faturamento */}
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.25rem', marginBottom: '1rem', background: 'rgba(255,255,255,0.02)', padding: '0.25rem', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                  {[
                    { id: 'particular', label: 'Particular' },
                    { id: 'unimed', label: 'Unimed' },
                    { id: 'cargill', label: 'Cargill' },
                    { id: 'chromatox', label: 'Chromatox' },
                    { id: 'desconsiderados', label: 'Ignorados' }
                  ].map(tab => (
                    <button
                      key={tab.id}
                      onClick={() => setFaturamentoSubTab(tab.id as any)}
                      className={`btn-filter ${faturamentoSubTab === tab.id ? 'active' : ''}`}
                      style={{
                        flex: 1,
                        padding: '0.35rem 0.5rem',
                        fontSize: '0.75rem',
                        borderRadius: '6px',
                        border: 'none',
                        background: faturamentoSubTab === tab.id ? 'var(--primary)' : 'transparent',
                        color: faturamentoSubTab === tab.id ? '#fff' : 'var(--text-muted)',
                        fontWeight: faturamentoSubTab === tab.id ? 600 : 400
                      }}
                    >
                      {tab.label}
                    </button>
                  ))}
                </div>

                {/* Busca */}
                <div className="search-container" style={{ marginBottom: '1rem' }}>
                  <Search size={16} className="search-icon" />
                  <input
                    type="text"
                    placeholder="Buscar por descrição ou valor..."
                    value={faturamentoBusca}
                    onChange={e => setFaturamentoBusca(e.target.value)}
                    className="search-input"
                    style={{ fontSize: '0.8rem', paddingLeft: '2.25rem' }}
                  />
                </div>

                {/* Lista de Transações */}
                <div className="table-responsive" style={{ maxHeight: '420px', overflowY: 'auto' }}>
                  <table className="table-conferencia" style={{ fontSize: '0.8rem' }}>
                    <thead>
                      <tr>
                        <th>Data</th>
                        <th>Descrição</th>
                        <th style={{ textAlign: 'right' }}>Valor</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredFaturamentoReceipts.map(r => (
                        <tr key={r.id}>
                          <td>{formatDate(r.data_lancamento)}</td>
                          <td style={{ whiteSpace: 'normal' }}>
                            <div style={{ fontWeight: 600, color: 'var(--text-dark)' }}>
                              {r.descricao_normalizada || 'S/D'}
                            </div>
                            <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)', marginTop: '0.15rem' }}>
                              {r.descricao_complementar}
                            </div>
                            {r.numero_documento && (
                              <div style={{ fontSize: '0.65rem', color: 'var(--primary)', marginTop: '0.1rem' }}>
                                Doc: {r.numero_documento}
                              </div>
                            )}
                          </td>
                          <td className="amount val-positive" style={{ textAlign: 'right', fontWeight: 600 }}>
                            {formatCurrency(r.valor)}
                          </td>
                        </tr>
                      ))}

                      {filteredFaturamentoReceipts.length === 0 && (
                        <tr>
                          <td colSpan={3} style={{ textAlign: 'center', padding: '3rem', color: 'var(--text-muted)' }}>
                            Nenhum lançamento encontrado nesta categoria.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>

                {/* Resumo da sub-aba */}
                <div style={{ marginTop: 'auto', paddingTop: '1rem', borderTop: '1px solid var(--border-color)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <span style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>
                    Total da Categoria:
                  </span>
                  <span style={{ fontSize: '1.1rem', fontWeight: 700, color: 'var(--success)' }}>
                    {formatCurrency(
                      faturamentoSubTab === 'particular' ? faturamentoData.totals.particular :
                      faturamentoSubTab === 'unimed' ? faturamentoData.totals.unimed :
                      faturamentoSubTab === 'cargill' ? faturamentoData.totals.cargill :
                      faturamentoSubTab === 'chromatox' ? faturamentoData.totals.chromatox :
                      faturamentoData.totals.disregarded
                    )}
                  </span>
                </div>
              </div>
            </div>

          </div>
        </div>
      )}

      {/* Tabs Content */}
      {activeTab === 'dashboard' && (
        <div className="fade-in">
          {/* Dinheiro Balance Summary Panel */}
          <div className="fade-in glass-card" style={{ marginBottom: '2rem', padding: '1.5rem', background: 'rgba(59, 130, 246, 0.03)', border: '1px solid rgba(59, 130, 246, 0.15)' }}>
            <h4 style={{ fontSize: '1rem', fontWeight: 600, marginBottom: '1.25rem', color: 'var(--primary)', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
              <DollarSign size={18} /> Resumo de Dinheiro em Caixa
            </h4>
            <div className="grid-4" style={{ gap: '1.5rem' }}>
              
              {/* Saldo Inicial */}
              <div>
                <div style={{ color: 'var(--text-muted)', fontSize: '0.85rem' }}>1. Saldo Inicial (Acumulado Anterior)</div>
                <div style={{ fontSize: '1.5rem', fontWeight: 700, color: 'var(--text-dark)', marginTop: '0.25rem' }}>
                  {formatCurrency(startBalance)}
                </div>
                <div style={{ color: 'var(--text-muted)', fontSize: '0.75rem', marginTop: '0.25rem' }}>
                  Saldo acumulado antes do período
                </div>
              </div>

              {/* Total Recebido PCLAB */}
              <div>
                <div style={{ color: 'var(--text-muted)', fontSize: '0.85rem' }}>2. (-) Faturamento em Dinheiro (PCLAB)</div>
                <div style={{ fontSize: '1.5rem', fontWeight: 700, color: 'var(--danger)', marginTop: '0.25rem' }}>
                  {formatCurrency(dinheiroTotalPclab)}
                </div>
                <div style={{ color: 'var(--text-muted)', fontSize: '0.75rem', marginTop: '0.25rem' }}>
                  Total faturado no período
                </div>
              </div>

              {/* Total Depositado BB */}
              <div>
                <div style={{ color: 'var(--text-muted)', fontSize: '0.85rem' }}>3. (+) Depósitos no Banco (Extrato)</div>
                <div style={{ fontSize: '1.5rem', fontWeight: 700, color: 'var(--success)', marginTop: '0.25rem' }}>
                  {formatCurrency(dinheiroTotalTarget)}
                </div>
                <div style={{ color: 'var(--text-muted)', fontSize: '0.75rem', marginTop: '0.25rem' }}>
                  Depósitos identificados no período
                </div>
              </div>

              {/* Saldo Final */}
              <div>
                <div style={{ color: 'var(--text-muted)', fontSize: '0.85rem' }}>4. (=) Saldo Final (Acumulado Caixa)</div>
                <div style={{ 
                  fontSize: '1.5rem', 
                  fontWeight: 700, 
                  color: (startBalance + dinheiroTotalTarget - dinheiroTotalPclab) >= 0 ? 'var(--success)' : 'var(--danger)',
                  marginTop: '0.25rem' 
                }}>
                  {formatCurrency(startBalance + dinheiroTotalTarget - dinheiroTotalPclab)}
                </div>
                <div style={{ color: 'var(--text-muted)', fontSize: '0.75rem', marginTop: '0.25rem' }}>
                  Saldo acumulado no caixa
                </div>
              </div>

            </div>
          </div>


          {/* Grids do Calendário e Detalhes */}
          <div className="dashboard-grid">
            <div className="col-span-8" style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>
              {calendarMonthsToRender.map(({ year, month }) => renderCalendar(year, month))}
              {calendarMonthsToRender.length === 0 && (
                <div className="glass-card" style={{ padding: '2rem', textAlign: 'center', color: 'var(--text-muted)' }}>
                  Selecione pelo menos um mês/ano nos filtros acima para visualizar o calendário do fluxo de caixa.
                </div>
              )}
            </div>

            <div className="col-span-4">
              {selectedCalendarDate ? (
                <div className="glass-card fade-in" style={{ position: 'sticky', top: '1rem', border: '1px solid var(--primary)', background: 'var(--bg-color-secondary)' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.25rem' }}>
                    <h3 style={{ color: 'var(--text-dark)', fontWeight: 600, fontSize: '1.1rem' }}>
                      Detalhes: {formatDate(selectedCalendarDate)}
                    </h3>
                    <button 
                      type="button" 
                      className="btn-filter" 
                      onClick={() => setSelectedCalendarDate(null)}
                      style={{ padding: '0.2rem 0.5rem', fontSize: '0.75rem', background: 'rgba(255,255,255,0.05)', borderRadius: '4px' }}
                    >
                      Fechar
                    </button>
                  </div>
                  
                  {(() => {
                    const stats = getCalendarDayStats(selectedCalendarDate);
                    return (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '0.35rem', marginBottom: '1.25rem', fontSize: '0.8rem', borderBottom: '1px solid var(--border-color)', paddingBottom: '0.75rem' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                          <span style={{ color: 'var(--text-muted)' }}>Faturamento do Dia:</span>
                          <span className="val-negative" style={{ fontWeight: 600 }}>-{formatCurrency(stats.faturamento)}</span>
                        </div>
                        <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                          <span style={{ color: 'var(--text-muted)' }}>Depósitos no Dia:</span>
                          <span className="val-positive" style={{ fontWeight: 600 }}>+{formatCurrency(stats.deposits)}</span>
                        </div>
                        <div style={{ display: 'flex', justifyContent: 'space-between', paddingTop: '0.25rem', borderTop: '1px solid rgba(255,255,255,0.03)' }}>
                          <span style={{ color: 'var(--text-dark)', fontWeight: 500 }}>Saldo Acumulado:</span>
                          <span className={stats.runningBalance >= 0 ? 'val-positive' : 'val-negative'} style={{ fontWeight: 700 }}>
                            {formatCurrency(stats.runningBalance)}
                          </span>
                        </div>
                      </div>
                    );
                  })()}

                  <div style={{ maxHeight: '350px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '0.5rem', paddingRight: '0.25rem' }}>
                    {selectedDayTransactions.length > 0 ? (
                      selectedDayTransactions.map((tx, idx) => (
                        <div 
                          key={tx.id || idx} 
                          style={{ 
                            padding: '0.6rem', 
                            borderRadius: '8px', 
                            background: 'rgba(255, 255, 255, 0.01)',
                            border: '1px solid var(--border-color)',
                            borderLeft: `4px solid ${tx.type === 'entrada' ? 'var(--success)' : 'var(--danger)'}`,
                            fontSize: '0.75rem'
                          }}
                        >
                          <div style={{ display: 'flex', justifyContent: 'space-between', fontWeight: 500, color: 'var(--text-dark)', gap: '0.5rem' }}>
                            <span style={{ whiteSpace: 'normal' }}>{tx.description}</span>
                            <span className={tx.type === 'entrada' ? 'val-positive' : 'val-negative'} style={{ whiteSpace: 'nowrap' }}>
                              {tx.type === 'entrada' ? '+' : '-'}{formatCurrency(tx.value)}
                            </span>
                          </div>
                          <div style={{ fontSize: '0.65rem', color: 'var(--text-muted)', marginTop: '0.25rem' }}>
                            Origem: {tx.source}
                          </div>
                        </div>
                      ))
                    ) : (
                      <div style={{ fontStyle: 'italic', color: 'var(--text-muted)', fontSize: '0.75rem', textAlign: 'center', padding: '2rem 1rem' }}>
                        Nenhum lançamento em dinheiro registrado neste dia.
                      </div>
                    )}
                  </div>
                </div>
              ) : (
                <div className="glass-card" style={{ position: 'sticky', top: '1rem', padding: '2rem 1.5rem', textAlign: 'center', color: 'var(--text-muted)', fontSize: '0.85rem' }}>
                  <Calendar size={28} style={{ color: 'var(--primary)', marginBottom: '0.75rem', display: 'block', margin: '0 auto 0.75rem' }} />
                  Clique em um dia do calendário para ver a lista de lançamentos em dinheiro do dia.
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {activeTab === 'conferencia' && (
        <div className="fade-in">
          {/* Summary KPIs (Apenas para Cartão e Pix/TED) */}
          <div className="dashboard-grid" style={{ marginBottom: '2rem' }}>
            <div className="glass-card col-span-4" style={{ display: 'flex', alignItems: 'center', gap: '1.25rem' }}>
              <div style={{ padding: '0.75rem', borderRadius: '12px', background: 'rgba(16, 185, 129, 0.1)', color: 'var(--success)' }}>
                <CheckCircle2 size={32} />
              </div>
              <div>
                <span className="kpi-title">Conciliados (Cartão e Pix)</span>
                <span className="kpi-value positive">
                  {formatCurrency(
                    reconciliationData.cartao.matched.reduce((acc, curr) => acc + curr.pclab.reduce((s, p) => s + Number(p.vr_lanc), 0), 0) +
                    reconciliationData.pix.matched.reduce((acc, curr) => acc + curr.pclab.reduce((s, p) => s + Number(p.vr_lanc), 0), 0)
                  )}
                </span>
                <div style={{ color: 'var(--text-muted)', fontSize: '0.8rem', marginTop: '0.25rem' }}>
                  {reconciliationData.cartao.matched.length + reconciliationData.pix.matched.length} grupos conciliados
                </div>
              </div>
            </div>

            <div className="glass-card col-span-4" style={{ display: 'flex', alignItems: 'center', gap: '1.25rem' }}>
              <div style={{ padding: '0.75rem', borderRadius: '12px', background: 'rgba(239, 68, 68, 0.1)', color: 'var(--danger)' }}>
                <ShieldAlert size={32} />
              </div>
              <div>
                <span className="kpi-title">Pendentes no PCLAB (Cartão e Pix)</span>
                <span className="kpi-value negative">
                  {formatCurrency(
                    reconciliationData.cartao.unmatchedPclab.reduce((acc, curr) => acc + Number(curr.vr_lanc), 0) +
                    reconciliationData.pix.unmatchedPclab.reduce((acc, curr) => acc + Number(curr.vr_lanc), 0)
                  )}
                </span>
                <div style={{ color: 'var(--text-muted)', fontSize: '0.8rem', marginTop: '0.25rem' }}>
                  {reconciliationData.cartao.unmatchedPclab.length + reconciliationData.pix.unmatchedPclab.length} faturamentos sem confirmação
                </div>
              </div>
            </div>

            <div className="glass-card col-span-4" style={{ display: 'flex', alignItems: 'center', gap: '1.25rem' }}>
              <div style={{ padding: '0.75rem', borderRadius: '12px', background: 'rgba(245, 158, 11, 0.1)', color: 'var(--warning)' }}>
                <Landmark size={32} />
              </div>
              <div>
                <span className="kpi-title">Pendentes no Destino (Cartão e Pix)</span>
                <span className="kpi-value" style={{ color: 'var(--warning)' }}>
                  {formatCurrency(
                    reconciliationData.cartao.unmatchedTarget.reduce((acc, curr) => acc + curr.valor, 0) +
                    reconciliationData.pix.unmatchedTarget.reduce((acc, curr) => acc + curr.valor, 0)
                  )}
                </span>
                <div style={{ color: 'var(--text-muted)', fontSize: '0.8rem', marginTop: '0.25rem' }}>
                  {reconciliationData.cartao.unmatchedTarget.length + reconciliationData.pix.unmatchedTarget.length} recebimentos sem faturamento
                </div>
              </div>
            </div>
          </div>

          {/* Sub Tab Controls */}
          <div className="glass-card" style={{ marginBottom: '2rem', padding: '1rem' }}>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '1rem', justifyContent: 'space-between', alignItems: 'center' }}>
              
              {/* Type Selectors */}
              <div className="button-group">
                <button className={`btn-filter ${confSubTab === 'cartao' ? 'active' : ''}`} onClick={() => setConfSubTab('cartao')}>
                  <CreditCard size={16} className="mr-2" /> Cartões
                </button>
                <button className={`btn-filter ${confSubTab === 'pix' ? 'active' : ''}`} onClick={() => setConfSubTab('pix')}>
                  <Banknote size={16} className="mr-2" /> Pix & Transferências
                </button>
              </div>

              {/* Excel and Filter options */}
              <div style={{ display: 'flex', gap: '1rem', alignItems: 'center', flexWrap: 'wrap' }}>
                {/* Search */}
                <div style={{ position: 'relative', minWidth: '240px' }}>
                  <input 
                    type="text" 
                    className="input-text" 
                    placeholder="Buscar por descrição ou valor..." 
                    value={confBusca} 
                    onChange={e => setConfBusca(e.target.value)} 
                    style={{ paddingLeft: '2.25rem' }} 
                  />
                  <Search size={16} style={{ position: 'absolute', left: '0.75rem', top: '50%', transform: 'translateY(-50%)', color: 'var(--text-muted)' }} />
                </div>

                {/* Date Filter */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                  <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>Data:</span>
                  <input 
                    type="date" 
                    className="input-text" 
                    value={confFiltroData} 
                    onChange={e => setConfFiltroData(e.target.value)} 
                    style={{ width: 'auto', padding: '0.4rem 0.75rem', background: 'var(--bg-card)', border: '1px solid var(--border-color)', color: 'var(--text-dark)', borderRadius: '6px' }} 
                  />
                  {confFiltroData && (
                    <button 
                      onClick={() => setConfFiltroData('')} 
                      className="btn-filter" 
                      style={{ padding: '0.4rem 0.6rem', fontSize: '0.75rem', borderRadius: '6px', border: '1px solid var(--border-color)' }}
                      title="Limpar filtro de data"
                    >
                      Limpar
                    </button>
                  )}
                </div>

                {/* Filter Selector */}
                <select className="input-select" value={confFiltroStatus} onChange={e => setConfFiltroStatus(e.target.value as any)} style={{ width: 'auto' }}>
                  <option value="todos">Exibir Todos os Status</option>
                  <option value="conciliados">Apenas Conciliados</option>
                  <option value="divergencias">Apenas Pendências / Divergências</option>
                </select>

                {/* Export Button */}
                <button onClick={() => handleExportConferencia(confSubTab)} className="btn-action" style={{ background: 'rgba(16, 185, 129, 0.1)', color: 'var(--success)', border: '1px solid var(--success)' }}>
                  <FileSpreadsheet size={16} className="mr-2" /> Exportar para Excel
                </button>
              </div>

            </div>
          </div>

          {/* Reconciliation Panels */}
          <div className="glass-card" style={{ padding: '1.5rem' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid var(--border-color)', paddingBottom: '0.75rem', marginBottom: '1.5rem' }}>
              <h3 style={{ fontSize: '1.1rem', fontWeight: 600 }}>
                Detalhamento de Conciliação: {confSubTab === 'cartao' ? 'Cartões' : 'Pix e Transferências'}
              </h3>
            </div>

            { (selectedPclabIds.length + selectedTargetIds.length >= 1) && (
              <div className="glass-card" style={{
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                padding: '1rem',
                marginBottom: '1rem',
                background: 'rgba(59, 130, 246, 0.1)',
                border: '1px solid rgba(59, 130, 246, 0.25)',
                borderRadius: '12px',
                animation: 'fadeIn 0.2s ease',
                position: 'sticky',
                top: 0,
                zIndex: 10
              }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
                  <ShieldAlert size={20} style={{ color: 'var(--primary)' }} />
                  <span style={{ fontSize: '0.9rem', fontWeight: 600 }}>
                    {selectedPclabIds.length} item(ns) PCLAB e {selectedTargetIds.length} item(ns) Destino selecionados para conciliação forçada.
                  </span>
                </div>
                <div style={{ display: 'flex', gap: '0.75rem' }}>
                  <button 
                    onClick={() => {
                      setSelectedPclabIds([]);
                      setSelectedTargetIds([]);
                    }}
                    className="btn-filter"
                    style={{ padding: '0.5rem 1rem', fontSize: '0.85rem' }}
                  >
                    Cancelar
                  </button>
                  <button 
                    onClick={handleForceReconcile}
                    className="btn-action"
                    style={{ padding: '0.5rem 1.25rem', fontSize: '0.85rem', background: 'var(--primary)', fontWeight: 600 }}
                  >
                    Confirmar Conciliação Forçada
                  </button>
                </div>
              </div>
            )}

            {/* Visualização Pareada para Cartão e Pix */}
            <div className="table-container">
              <table style={{ tableLayout: 'auto', width: '100%' }}>
                <thead>
                  <tr>
                    <th style={{ width: '40px', textAlign: 'center' }}></th>
                    <th style={{ width: '105px' }}>Status</th>
                    <th>PCLAB Lançamento(s)</th>
                    <th className="amount" style={{ width: '95px' }}>Valor PCLAB</th>
                    <th style={{ textAlign: 'center', width: '35px' }}>Conf.</th>
                    <th>Destino (Maquininha/Banco)</th>
                    <th className="amount" style={{ width: '95px' }}>Valor Destino</th>
                    <th className="amount" style={{ width: '85px' }}>Diferença</th>
                  </tr>
                </thead>
                <tbody>
                  {sortedReconciliationRows.map((row, idx) => {
                    if (row.type === 'matched') {
                      const item = row.data;
                      return (
                        <tr key={`matched-${idx}`} className="bg-row-success">
                          <td style={{ textAlign: 'center' }}></td>
                          <td>
                            <span className={`badge ${item.isForced ? 'badge-blue' : item.isCrossDate ? 'badge-yellow' : item.isGroup ? 'badge-blue' : 'badge-green'}`} style={{ whiteSpace: 'normal', fontSize: '0.7rem', padding: '0.2rem 0.4rem', textAlign: 'center', display: 'block' }}>
                              {item.isForced ? 'Conciliado (Forçado)' : item.isCrossDate ? 'Conciliado (Data Dif.)' : item.isGroup ? 'Conciliado (Lote)' : 'Conciliado'}
                            </span>
                          </td>
                          <td style={{ whiteSpace: 'normal', wordBreak: 'break-word' }}>
                            {item.pclab.map((p, pIdx) => (
                              <div key={p.id || pIdx} style={{ marginBottom: pIdx < item.pclab.length - 1 ? '0.4rem' : 0 }}>
                                <div style={{ fontWeight: 500 }}>{p.descricao}</div>
                                <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)' }}>
                                  {formatDateHour(p.dt_lancamento)} | {p.forma_pgto}
                                </div>
                              </div>
                            ))}
                          </td>
                          <td className="amount val-positive">
                            {item.pclab.length === 1 ? (
                              formatCurrency(item.pclab[0].vr_lanc)
                            ) : (
                              <div>
                                <div style={{ fontWeight: 600 }}>{formatCurrency(item.pclab.reduce((a, b) => a + Number(b.vr_lanc), 0))}</div>
                                <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)', whiteSpace: 'normal', maxWidth: '150px' }}>
                                  {item.pclab.map(p => formatCurrency(p.vr_lanc)).join(' + ')}
                                </div>
                              </div>
                            )}
                          </td>
                          <td style={{ textAlign: 'center', color: 'var(--success)', padding: '1rem 0.2rem' }}><ArrowRight size={16} style={{ display: 'inline' }} /></td>
                          <td style={{ whiteSpace: 'normal', wordBreak: 'break-word' }}>
                            {item.target.map((t, tIdx) => (
                              <div key={t.id || tIdx} style={{ marginBottom: tIdx < item.target.length - 1 ? '0.4rem' : 0 }}>
                                <div style={{ fontWeight: 500 }}>{t.descricao}</div>
                                <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)' }}>
                                  {formatDateHour(t.data)} | {t.tipo}
                                </div>
                              </div>
                            ))}
                          </td>
                          <td className="amount val-positive">
                            {item.target.length === 1 ? (
                              formatCurrency(item.target[0].valor)
                            ) : (
                              <div>
                                <div style={{ fontWeight: 600 }}>{formatCurrency(item.target.reduce((a, b) => a + Number(b.valor), 0))}</div>
                                <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)', whiteSpace: 'normal', maxWidth: '150px' }}>
                                  {item.target.map(t => formatCurrency(t.valor)).join(' + ')}
                                </div>
                              </div>
                            )}
                          </td>
                          <td className="amount" style={{ color: 'var(--success)', fontWeight: 600 }}>R$ 0,00</td>
                        </tr>
                      );
                    } else if (row.type === 'pclab') {
                      const item = row.data;
                      return (
                        <tr key={`unmatched-pclab-${idx}`} className="bg-row-danger">
                          <td style={{ textAlign: 'center' }}>
                            <input 
                              type="checkbox"
                              checked={selectedPclabIds.includes(item.id)}
                              onChange={() => handleTogglePclab(item.id)}
                              style={{ width: '16px', height: '16px', accentColor: 'var(--primary)', cursor: 'pointer' }}
                            />
                          </td>
                          <td>
                            <span className="badge badge-red" style={{ whiteSpace: 'normal', fontSize: '0.7rem', padding: '0.2rem 0.4rem', textAlign: 'center', display: 'block' }}>
                              Pendente PCLAB
                            </span>
                          </td>
                          <td style={{ whiteSpace: 'normal', wordBreak: 'break-word' }}>
                            <div style={{ fontWeight: 500 }}>{item.descricao}</div>
                            <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)' }}>{formatDateHour(item.dt_lancamento)} | {item.forma_pgto}</div>
                          </td>
                          <td className="amount val-negative">{formatCurrency(item.vr_lanc)}</td>
                          <td style={{ textAlign: 'center', color: 'var(--danger)', padding: '1rem 0.2rem' }}>-</td>
                          <td style={{ color: 'var(--text-muted)', fontStyle: 'italic', whiteSpace: 'normal', wordBreak: 'break-word' }}>Não encontrado no destino (Maquininha/Extrato)</td>
                          <td className="amount">R$ 0,00</td>
                          <td className="amount val-negative">{formatCurrency(item.vr_lanc)}</td>
                        </tr>
                      );
                    } else {
                      const item = row.data;
                      return (
                        <tr key={`unmatched-target-${idx}`} className="bg-row-danger" style={{ backgroundColor: 'rgba(245, 158, 11, 0.06)' }}>
                          <td style={{ textAlign: 'center' }}>
                            <input 
                              type="checkbox"
                              checked={selectedTargetIds.includes(item.id)}
                              onChange={() => handleToggleTarget(item.id)}
                              style={{ width: '16px', height: '16px', accentColor: 'var(--primary)', cursor: 'pointer' }}
                            />
                          </td>
                          <td>
                            <span className="badge badge-yellow" style={{ whiteSpace: 'normal', fontSize: '0.7rem', padding: '0.2rem 0.4rem', textAlign: 'center', display: 'block' }}>
                              Pendente Destino
                            </span>
                          </td>
                          <td style={{ color: 'var(--text-muted)', fontStyle: 'italic', whiteSpace: 'normal', wordBreak: 'break-word' }}>Não registrado no faturamento (PCLAB)</td>
                          <td className="amount">R$ 0,00</td>
                          <td style={{ textAlign: 'center', color: 'var(--warning)', padding: '1rem 0.2rem' }}>-</td>
                          <td style={{ whiteSpace: 'normal', wordBreak: 'break-word' }}>
                            <div style={{ fontWeight: 500 }}>{item.descricao}</div>
                            <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)' }}>
                              {formatDateHour(item.data)} | {item.tipo}
                            </div>
                          </td>
                          <td className="amount" style={{ color: 'var(--warning)' }}>{formatCurrency(item.valor)}</td>
                          <td className="amount" style={{ color: 'var(--warning)' }}>{formatCurrency(item.valor)}</td>
                        </tr>
                      );
                    }
                  })}

                  {/* Empty state */}
                  {sortedReconciliationRows.length === 0 && (
                    <tr>
                      <td colSpan={8} style={{ textAlign: 'center', padding: '3rem', color: 'var(--text-muted)' }}>
                        Nenhum registro encontrado com os filtros atuais.
                      </td>
                    </tr>
                  )}
                  </tbody>
                </table>
              </div>
            </div>
        </div>
      )}

      {activeTab === 'classificacao' && (
        <div className="fade-in glass-card">
          <h2 style={{ fontSize: '1.25rem', fontWeight: 600, marginBottom: '0.5rem' }}>Classificação de Receitas Bancárias</h2>
          <p style={{ color: 'var(--text-muted)', fontSize: '0.875rem', marginBottom: '2rem' }}>
            Listagem de receitas recebidas na conta do Banco do Brasil que ainda não possuem classificação de regras definida.
          </p>

          {unclassifiedGroups.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '4rem', color: 'var(--success)' }}>
              <BadgeCheck size={48} style={{ display: 'block', margin: '0 auto 1rem auto' }} />
              <div style={{ fontWeight: 600, fontSize: '1.1rem' }}>Todas as receitas estão classificadas!</div>
              <p style={{ color: 'var(--text-muted)', fontSize: '0.85rem', marginTop: '0.25rem' }}>Nenhum lançamento pendente encontrado no extrato bancário.</p>
            </div>
          ) : (
            <div className="table-container">
              <table>
                <thead>
                  <tr>
                    <th>Descrição no Extrato</th>
                    <th className="amount">Lançamentos</th>
                    <th className="amount">Valor Total</th>
                    <th>Último Lançamento</th>
                    <th style={{ textAlign: 'center', width: '320px' }}>Classificar Como</th>
                  </tr>
                </thead>
                <tbody>
                  {unclassifiedGroups.map((group, idx) => (
                    <tr key={`unclass-${idx}`}>
                      <td style={{ fontWeight: 600, whiteSpace: 'normal', maxWidth: '300px' }}>{group.descricao_normalizada}</td>
                      <td className="amount" style={{ fontWeight: 500 }}>{group.quantidade}x</td>
                      <td className="amount val-positive" style={{ fontWeight: 600 }}>{formatCurrency(group.total_valor)}</td>
                      <td>
                        <div style={{ fontWeight: 500 }}>{formatDate(group.data_ultimo)}</div>
                        <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)' }}>Último valor: {formatCurrency(group.valor_ultimo)}</div>
                      </td>
                      <td>
                        <div className="grid-actions" style={{ gap: '0.25rem' }}>
                          <button 
                            disabled={savingClass === group.descricao_normalizada} 
                            onClick={() => handleClassificar(group.descricao_normalizada, 'Dinheiro')} 
                            className="btn-action" 
                            style={{ padding: '0.4rem 0.25rem', fontSize: '0.75rem', background: 'rgba(16, 185, 129, 0.1)', color: 'var(--success)' }}
                          >
                            Dinheiro
                          </button>
                          <button 
                            disabled={savingClass === group.descricao_normalizada} 
                            onClick={() => handleClassificar(group.descricao_normalizada, 'Transferencia Bancária')} 
                            className="btn-action" 
                            style={{ padding: '0.4rem 0.25rem', fontSize: '0.75rem', background: 'rgba(59, 130, 246, 0.1)', color: 'var(--primary)' }}
                          >
                            Transf
                          </button>
                          <button 
                            disabled={savingClass === group.descricao_normalizada} 
                            onClick={() => handleClassificar(group.descricao_normalizada, 'Cartão')} 
                            className="btn-action" 
                            style={{ padding: '0.4rem 0.25rem', fontSize: '0.75rem', background: 'rgba(245, 158, 11, 0.1)', color: 'var(--warning)' }}
                          >
                            Cartão
                          </button>
                          <button 
                            disabled={savingClass === group.descricao_normalizada} 
                            onClick={() => handleClassificar(group.descricao_normalizada, 'desconsiderar')} 
                            className="btn-action" 
                            style={{ padding: '0.4rem 0.25rem', fontSize: '0.75rem', background: 'rgba(239, 68, 68, 0.1)', color: 'var(--danger)' }}
                          >
                            Ignorar
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {activeTab === 'regras' && (
        <div className="fade-in glass-card">
          <h2 style={{ fontSize: '1.25rem', fontWeight: 600, marginBottom: '0.5rem' }}>Regras de Classificação Cadastradas</h2>
          <p style={{ color: 'var(--text-muted)', fontSize: '0.875rem', marginBottom: '2rem' }}>
            Gerencie as regras automáticas criadas para classificar as descrições do extrato bancário.
          </p>

          {regrasClassificacao.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '4rem', color: 'var(--text-muted)' }}>
              Nenhuma regra de classificação encontrada.
            </div>
          ) : (
            <div className="table-container">
              <table>
                <thead>
                  <tr>
                    <th>Descrição Normalizada</th>
                    <th>Classificação Associada</th>
                    <th>Cadastrado em</th>
                    <th style={{ textAlign: 'center', width: '80px' }}>Ações</th>
                  </tr>
                </thead>
                <tbody>
                  {regrasClassificacao.map((regra, idx) => (
                    <tr key={`regra-${idx}`}>
                      <td style={{ fontWeight: 500, whiteSpace: 'normal' }}>{regra.descricao_normalizada}</td>
                      <td>
                        <span className={`badge ${
                          regra.classificacao === 'Dinheiro' ? 'badge-green' : 
                          regra.classificacao === 'Transferencia Bancária' ? 'badge-blue' : 
                          regra.classificacao === 'Cartão' ? 'badge-yellow' : 'badge-red'
                        }`}>
                          {regra.classificacao === 'desconsiderar' ? 'Desconsiderar / Ignorar' : regra.classificacao}
                        </span>
                      </td>
                      <td>{formatDate(regra.criado_em)}</td>
                      <td style={{ textAlign: 'center' }}>
                        <button onClick={() => handleDeleteRegra(regra.descricao_normalizada)} className="btn-filter" style={{ padding: '0.25rem', color: 'var(--danger)' }}>
                          <Trash2 size={16} />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {activeTab === 'ajustes' && (
        <div className="fade-in grid-split" style={{ gap: '2rem', alignItems: 'start' }}>
          {/* Card Esquerdo: Lançar Ajuste */}
          <div className="glass-card" style={{ padding: '1.5rem' }}>
            <h3 style={{ fontSize: '1.1rem', fontWeight: 600, marginBottom: '1.25rem', display: 'flex', alignItems: 'center', gap: '0.5rem', color: 'var(--primary)' }}>
              <ArrowRightLeft size={18} style={{ transform: 'rotate(90deg)' }} /> Lançar Ajuste Manual
            </h3>
            <p style={{ color: 'var(--text-muted)', fontSize: '0.8rem', marginBottom: '1.5rem' }}>
              Crie registros virtuais para compensar pendências e reconciliar os saldos da conciliação.
            </p>
            <form onSubmit={handleCreateAjuste} style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>
              <div>
                <label style={{ display: 'block', fontSize: '0.8rem', fontWeight: 500, marginBottom: '0.35rem', color: 'var(--text-dark)' }}>Data do Ajuste</label>
                <input 
                  type="date" 
                  className="input-text" 
                  value={ajustFormDate} 
                  onChange={e => setAjustFormDate(e.target.value)} 
                  required
                />
              </div>

              <div>
                <label style={{ display: 'block', fontSize: '0.8rem', fontWeight: 500, marginBottom: '0.35rem', color: 'var(--text-dark)' }}>Categoria da Conciliação</label>
                <select 
                  className="input-select" 
                  value={ajustFormCat} 
                  onChange={e => setAjustFormCat(e.target.value)}
                >
                  <option value="cartao">Cartões</option>
                  <option value="pix">Pix & Transferências</option>
                  <option value="dinheiro">Dinheiro (Saldo)</option>
                </select>
              </div>

              <div>
                <label style={{ display: 'block', fontSize: '0.8rem', fontWeight: 500, marginBottom: '0.35rem', color: 'var(--text-dark)' }}>Lado do Lançamento</label>
                <select 
                  className="input-select" 
                  value={ajustFormLado} 
                  onChange={e => setAjustFormLado(e.target.value)}
                >
                  <option value="pclab">PCLAB (Faturamento)</option>
                  <option value="destino">Destino (Extrato/Maquininha)</option>
                </select>
              </div>

              <div>
                <label style={{ display: 'block', fontSize: '0.8rem', fontWeight: 500, marginBottom: '0.35rem', color: 'var(--text-dark)' }}>Descrição</label>
                <input 
                  type="text" 
                  className="input-text" 
                  placeholder="Ex: Ajuste faturamento não registrado" 
                  value={ajustFormDesc} 
                  onChange={e => setAjustFormDesc(e.target.value)} 
                  required
                />
              </div>

              <div>
                <label style={{ display: 'block', fontSize: '0.8rem', fontWeight: 500, marginBottom: '0.35rem', color: 'var(--text-dark)' }}>Valor (R$)</label>
                <input 
                  type="number" 
                  step="0.01" 
                  className="input-text" 
                  placeholder="0.00" 
                  value={ajustFormVal} 
                  onChange={e => setAjustFormVal(e.target.value)} 
                  required
                />
                <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)', display: 'block', marginTop: '0.25rem' }}>
                  Use valores positivos (ex: 150) ou negativos (ex: -150) para somar ou subtrair do saldo.
                </span>
              </div>

              <button 
                type="submit" 
                className="btn-action" 
                style={{ width: '100%', marginTop: '0.5rem', background: 'var(--primary)', color: '#fff', display: 'flex', justifyContent: 'center' }}
                disabled={savingAjust}
              >
                {savingAjust ? 'Registrando...' : 'Gravar Ajuste'}
              </button>
            </form>
          </div>

          {/* Card Direito: Tabela de Ajustes no Período */}
          <div className="glass-card" style={{ padding: '1.5rem' }}>
            <h3 style={{ fontSize: '1.1rem', fontWeight: 600, marginBottom: '1.25rem', color: 'var(--text-dark)' }}>
              Ajustes Registrados no Período
            </h3>
            <p style={{ color: 'var(--text-muted)', fontSize: '0.8rem', marginBottom: '1.5rem' }}>
              Estes ajustes são injetados como lançamentos virtuais na conciliação selecionada.
            </p>

            <div className="table-container" style={{ maxHeight: '500px', overflowY: 'auto' }}>
              <table>
                <thead>
                  <tr>
                    <th>Data</th>
                    <th>Categoria</th>
                    <th>Lado</th>
                    <th>Descrição</th>
                    <th className="amount">Valor</th>
                    <th style={{ textAlign: 'center', width: '80px' }}>Ações</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredAjustes.map((ajuste, idx) => (
                    <tr key={`ajuste-row-${idx}`}>
                      <td>{formatDate(ajuste.data)}</td>
                      <td>
                        <span className="badge" style={{ background: 'rgba(59, 130, 246, 0.1)', color: 'var(--primary)' }}>
                          {ajuste.categoria === 'cartao' ? 'Cartões' : ajuste.categoria === 'pix' ? 'Pix' : 'Dinheiro'}
                        </span>
                      </td>
                      <td>
                        <span className={`badge ${ajuste.lado === 'pclab' ? 'badge-blue' : 'badge-green'}`}>
                          {ajuste.lado === 'pclab' ? 'PCLAB' : 'Destino'}
                        </span>
                      </td>
                      <td style={{ whiteSpace: 'normal', fontSize: '0.8rem' }}>{ajuste.descricao}</td>
                      <td className={`amount ${ajuste.valor >= 0 ? 'val-positive' : 'val-negative'}`} style={{ fontWeight: 600 }}>
                        {formatCurrency(ajuste.valor)}
                      </td>
                      <td style={{ textAlign: 'center' }}>
                        <button onClick={() => handleDeleteAjuste(ajuste.id)} className="btn-filter" style={{ padding: '0.25rem', color: 'var(--danger)' }}>
                          <Trash2 size={16} />
                        </button>
                      </td>
                    </tr>
                  ))}

                  {filteredAjustes.length === 0 && (
                    <tr>
                      <td colSpan={6} style={{ textAlign: 'center', padding: '3rem', color: 'var(--text-muted)' }}>
                        Nenhum ajuste manual registrado para o período filtrado.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
