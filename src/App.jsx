import React, { useState, useEffect, useMemo, useRef } from 'react';
import { Wallet, ListChecks, LayoutGrid, Plus, Trash2, Pencil, X, ChevronDown, ChevronUp, Paperclip, Download, ShoppingCart, Link2, Check, MessageSquareWarning, PiggyBank, FolderOpen } from 'lucide-react';
import * as XLSX from 'xlsx';
import { loadSharedState, saveSharedState, subscribeToSharedState } from './lib/sharedStorage.js';

const STAGE_INFO = [
  { name: 'Этап 1', payment: 3406800 },
  { name: 'Этап 2', payment: 4542400 },
  { name: 'Этап 3', payment: 2839000 },
  { name: 'Финальный расчёт', payment: 567800 },
];
const STAGES = STAGE_INFO.map((s) => s.name);
const STAGE_PAYMENT = Object.fromEntries(STAGE_INFO.map((s) => [s.name, s.payment]));

function buildDefaultTasks() {
  return [
    // Этап 1 — предоплата 3 406 800 ₸
    'Демонтажные работы', 'Монтаж ютп электрокабеля', 'Грунтовка стен', 'Установка маяков',
    'Штукатурка стен (гипсовая штукатурка ротбанд)', 'Монтаж гкл перегородок в четыре слоя',
    'Подготовительные работы', 'Вынос мусора', 'Демонтаж пластиковой двери',
  ].map((title) => ({ title, stage: 'Этап 1' }))
    .concat([
      // Этап 2 — предоплата 4 542 400 ₸
      'Установка электро щитка', 'Грунтовка пола', 'Монтаж сантехнических труб (лучевая система)',
      'Левкас стен', 'Монтаж гкл коробов и стен', 'Шкурка стен', 'Грунтовка подготовка пола',
      'Установка подрозетников', 'Заливка наливного пола', 'Гидроизоляция сан узлов',
    ].map((title) => ({ title, stage: 'Этап 2' })))
    .concat([
      // Этап 3 — предоплата 2 839 000 ₸
      'Укладка кафеля (сан узлы, кухня пол)', 'Монтаж пола (ламинат обычный)', 'Покраска водоэмульсией (обои)',
      'Монтаж настенных бра', 'Монтаж розеток, вкл', 'Затирка швов кафеля', 'Монтаж плинтуса',
      'Установка сан фаянсов (унитаз, ванна, смеситель, душевая лейка)', 'Монтаж бойлера', 'Пред чистовая уборка',
    ].map((title) => ({ title, stage: 'Этап 3' })))
    .concat([{ title: 'Финальная приёмка и сдача работ (окончательный расчёт)', stage: 'Финальный расчёт' }])
    .map((t) => ({ id: uid('t'), title: t.title, roomId: '', stage: t.stage, cost: 0, done: false }));
}

const PURCHASE_TYPES = ['Черновые материалы', 'Чистовые материалы', 'Расходники', 'Техника', 'Мебель готовая', 'Мебель на заказ', 'Сантехника', 'Электрика', 'Двери', 'Потолок'];
const PURCHASE_STATUSES = ['В плане', 'Есть продавец', 'Заказано', 'Куплено'];
const PURCHASE_STATUS_META = {
  'В плане': { label: 'В плане', color: '#8B8578', bg: '#EFEDE6' },
  'Есть продавец': { label: 'Есть продавец', color: '#B8860B', bg: '#FBF0D9' },
  'Заказано': { label: 'Заказано', color: '#3B6EA5', bg: '#E1EAF5' },
  'Куплено': { label: 'Куплено', color: '#3F7D58', bg: '#E1EEE4' },
};

const EXPENSE_SUBTYPES = ['Работа', 'Материалы'];

const ISSUE_STATUSES = ['В планах', 'Отложено', 'В работе', 'Завершено'];
const RESPONSIBLE_PEOPLE = ['Юля', 'Коля', 'Александр'];
const ISSUE_STATUS_META = {
  'В планах': { color: '#8B8578', bg: '#EFEDE6' },
  'Отложено': { color: '#B8860B', bg: '#FBF0D9' },
  'В работе': { color: '#3B6EA5', bg: '#E1EAF5' },
  'Завершено': { color: '#3F7D58', bg: '#E1EEE4' },
};

const ADVANCE_STATUSES = ['Расходуется', 'Израсходован'];
const ADVANCE_STATUS_META = {
  'Расходуется': { color: '#B8860B', bg: '#FBF0D9' },
  'Израсходован': { color: '#3F7D58', bg: '#E1EEE4' },
};

function formatDateRu(dateStr) {
  if (!dateStr) return '—';
  try {
    const d = new Date(dateStr + 'T00:00:00');
    if (isNaN(d.getTime())) return dateStr;
    return new Intl.DateTimeFormat('ru-RU', { day: 'numeric', month: 'long', year: 'numeric' }).format(d);
  } catch (e) {
    return dateStr;
  }
}

const DEFAULT_ROOMS = [
  { id: 'r1', name: 'Прихожая', area: 0, planned: 0 },
  { id: 'r2', name: 'Кухня', area: 0, planned: 0 },
  { id: 'r3', name: 'Кабинет', area: 0, planned: 0 },
  { id: 'r4', name: 'Гостиная', area: 0, planned: 0 },
  { id: 'r5', name: 'Коридор общий', area: 0, planned: 0 },
  { id: 'r6', name: 'Спальня Гостевая', area: 0, planned: 0 },
  { id: 'r7', name: 'Мастер-спальня', area: 0, planned: 0 },
  { id: 'r8', name: 'Санузел Детский', area: 0, planned: 0 },
  { id: 'r9', name: 'Санузел Мастер', area: 0, planned: 0 },
  { id: 'r10', name: 'Санузел Прихожая', area: 0, planned: 0 },
  { id: 'r11', name: 'Коридор детский', area: 0, planned: 0 },
  { id: 'r12', name: 'Детская Сава', area: 0, planned: 0 },
  { id: 'r13', name: 'Детская Лева', area: 0, planned: 0 },
  { id: 'r14', name: 'Гардеробная Детская', area: 0, planned: 0 },
  { id: 'r15', name: 'Прачечная', area: 0, planned: 0 },
];

const uid = (p = 'id') => p + '_' + Math.random().toString(36).slice(2, 9);
const fmt = (n) => new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 0 }).format(Math.round(n || 0)) + ' ₸';

const COLORS = {
  bg: '#F5F3EE',
  surface: '#FFFFFF',
  ink: '#1E2A3A',
  inkSoft: '#5C6B7A',
  accent: '#E2572B',
  accentSoft: '#FBE4DA',
  line: '#DAD5C9',
  success: '#3F7D58',
};

function BlueprintCard({ children, style }) {
  return (
    <div
      style={{
        position: 'relative',
        background: COLORS.surface,
        border: `1px solid ${COLORS.line}`,
        borderRadius: 10,
        padding: 16,
        ...style,
      }}
    >
      {['tl', 'tr', 'bl', 'br'].map((pos) => (
        <span
          key={pos}
          style={{
            position: 'absolute',
            width: 8,
            height: 8,
            borderColor: COLORS.accent,
            opacity: 0.55,
            [pos.includes('t') ? 'top' : 'bottom']: -1,
            [pos.includes('l') ? 'left' : 'right']: -1,
            borderTopWidth: pos.includes('t') ? 2 : 0,
            borderBottomWidth: pos.includes('b') ? 2 : 0,
            borderLeftWidth: pos.includes('l') ? 2 : 0,
            borderRightWidth: pos.includes('r') ? 2 : 0,
            borderStyle: 'solid',
          }}
        />
      ))}
      {children}
    </div>
  );
}

function ProgressBar({ pct, color = COLORS.accent, track = COLORS.line, height = 8 }) {
  return (
    <div style={{ width: '100%', height, background: track, borderRadius: 999, overflow: 'hidden' }}>
      <div
        style={{
          width: `${Math.min(100, Math.max(0, pct))}%`,
          height: '100%',
          background: color,
          borderRadius: 999,
          transition: 'width 0.4s ease',
        }}
      />
    </div>
  );
}

function StatCard({ label, value, sub }) {
  return (
    <BlueprintCard style={{ flex: '1 1 160px', minWidth: 160 }}>
      <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 11, letterSpacing: '0.06em', textTransform: 'uppercase', color: COLORS.inkSoft, marginBottom: 8 }}>
        {label}
      </div>
      <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 24, fontWeight: 700, color: COLORS.ink, fontVariantNumeric: 'tabular-nums' }}>
        {value}
      </div>
      {sub && <div style={{ fontFamily: 'IBM Plex Mono, monospace', fontSize: 12, color: COLORS.inkSoft, marginTop: 4 }}>{sub}</div>}
    </BlueprintCard>
  );
}

function IconBtn({ onClick, title, children, danger }) {
  return (
    <button
      onClick={onClick}
      title={title}
      style={{
        display: 'inline-flex',
        alignItems: 'center',
        justifyContent: 'center',
        width: 30,
        height: 30,
        borderRadius: 7,
        border: `1px solid ${COLORS.line}`,
        background: COLORS.surface,
        color: danger ? '#B23A2C' : COLORS.inkSoft,
        cursor: 'pointer',
      }}
    >
      {children}
    </button>
  );
}

function TextInput(props) {
  return (
    <input
      {...props}
      style={{
        fontFamily: 'Inter, sans-serif',
        fontSize: 14,
        padding: '8px 10px',
        borderRadius: 7,
        border: `1px solid ${COLORS.line}`,
        background: COLORS.surface,
        color: COLORS.ink,
        outline: 'none',
        width: '100%',
        boxSizing: 'border-box',
        ...(props.style || {}),
      }}
    />
  );
}

function TextArea(props) {
  const { maxLength = 200, value, ...rest } = props;
  return (
    <div>
      <textarea
        {...rest}
        value={value}
        maxLength={maxLength}
        rows={props.rows || 3}
        style={{
          fontFamily: 'Inter, sans-serif',
          fontSize: 14,
          padding: '8px 10px',
          borderRadius: 7,
          border: `1px solid ${COLORS.line}`,
          background: COLORS.surface,
          color: COLORS.ink,
          outline: 'none',
          width: '100%',
          boxSizing: 'border-box',
          resize: 'vertical',
          ...(props.style || {}),
        }}
      />
      <div style={{ textAlign: 'right', fontFamily: 'IBM Plex Mono, monospace', fontSize: 10, color: COLORS.inkSoft, marginTop: 2 }}>
        {(value || '').length}/{maxLength}
      </div>
    </div>
  );
}

function Select(props) {
  return (
    <select
      {...props}
      style={{
        fontFamily: 'Inter, sans-serif',
        fontSize: 14,
        padding: '8px 10px',
        borderRadius: 7,
        border: `1px solid ${COLORS.line}`,
        background: COLORS.surface,
        color: COLORS.ink,
        outline: 'none',
        width: '100%',
        boxSizing: 'border-box',
        ...(props.style || {}),
      }}
    >
      {props.children}
    </select>
  );
}

export default function RemontTracker() {
  const [rooms, setRooms] = useState(DEFAULT_ROOMS);
  const [tasks, setTasks] = useState(() => buildDefaultTasks());
  const [expenses, setExpenses] = useState([]);
  const [purchases, setPurchases] = useState([]);
  const [issues, setIssues] = useState([]);
  const [advances, setAdvances] = useState([]);
  const [documents, setDocuments] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [tab, setTab] = useState('overview');

  // Set by applyRemoteData right before a batch of setState calls sourced from
  // Supabase (initial load or another collaborator's realtime update), so the
  // save-effect below can skip re-uploading data that just came from the server.
  const isRemoteUpdate = useRef(false);
  const saveTimeout = useRef(null);

  const applyRemoteData = (data) => {
    isRemoteUpdate.current = true;
    setRooms(data.rooms && data.rooms.length ? data.rooms : DEFAULT_ROOMS);
    setTasks(data.tasks && data.tasks.length ? data.tasks : buildDefaultTasks());
    setExpenses(data.expenses || []);
    setPurchases(data.purchases || []);
    setIssues(data.issues || []);
    setAdvances(data.advances || []);
    setDocuments(data.documents || []);
  };

  useEffect(() => {
    let unsubscribe = () => {};
    (async () => {
      const defaultData = {
        rooms: DEFAULT_ROOMS,
        tasks: buildDefaultTasks(),
        expenses: [],
        purchases: [],
        issues: [],
        advances: [],
        documents: [],
      };
      const data = await loadSharedState(defaultData);
      applyRemoteData(data);
      setLoaded(true);
      unsubscribe = subscribeToSharedState((remoteData) => applyRemoteData(remoteData));
    })();
    return () => unsubscribe();
  }, []);

  useEffect(() => {
    if (!loaded) return;
    if (isRemoteUpdate.current) {
      isRemoteUpdate.current = false;
      return;
    }
    if (saveTimeout.current) clearTimeout(saveTimeout.current);
    saveTimeout.current = setTimeout(() => {
      saveSharedState({ rooms, tasks, expenses, purchases, issues, advances, documents });
    }, 600);
    return () => clearTimeout(saveTimeout.current);
  }, [rooms, tasks, expenses, purchases, issues, advances, documents, loaded]);

  const totalPlanned = useMemo(() => rooms.reduce((s, r) => s + (Number(r.planned) || 0), 0), [rooms]);
  const totalSpent = useMemo(() => expenses.reduce((s, e) => s + (Number(e.amount) || 0), 0), [expenses]);
  const totalArea = useMemo(() => rooms.reduce((s, r) => s + (Number(r.area) || 0), 0), [rooms]);
  const tasksDone = tasks.filter((t) => t.done).length;
  const overallProgress = tasks.length ? Math.round((tasksDone / tasks.length) * 100) : 0;

  const spentByRoom = useMemo(() => {
    const map = {};
    expenses.forEach((e) => { map[e.roomId] = (map[e.roomId] || 0) + Number(e.amount || 0); });
    return map;
  }, [expenses]);

  const tasksByRoom = useMemo(() => {
    const map = {};
    tasks.forEach((t) => {
      if (!map[t.roomId]) map[t.roomId] = [];
      map[t.roomId].push(t);
    });
    return map;
  }, [tasks]);

  if (!loaded) {
    return (
      <div style={{ padding: 40, textAlign: 'center', fontFamily: 'Inter, sans-serif', color: COLORS.inkSoft }}>
        Загрузка проекта…
      </div>
    );
  }

  const tabs = [
    { id: 'overview', label: 'Обзор', icon: LayoutGrid },
    { id: 'budget', label: 'Смета', icon: Wallet },
    { id: 'advances', label: 'Авансы', icon: PiggyBank },
    { id: 'plan', label: 'План работ', icon: ListChecks },
    { id: 'purchases', label: 'Покупки', icon: ShoppingCart },
    { id: 'issues', label: 'Вопросы', icon: MessageSquareWarning },
    { id: 'documents', label: 'Проект', icon: FolderOpen },
  ];

  return (
    <div style={{ background: COLORS.bg, minHeight: '100%', padding: '20px 16px 40px', boxSizing: 'border-box' }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; }
        button { font-family: inherit; }
        button:focus-visible, input:focus-visible, select:focus-visible {
          outline: 2px solid ${COLORS.accent};
          outline-offset: 1px;
        }
      `}</style>

      <div style={{ maxWidth: 880, margin: '0 auto' }}>
        <header style={{ marginBottom: 20 }}>
          <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 12, letterSpacing: '0.08em', textTransform: 'uppercase', color: COLORS.accent, fontWeight: 600 }}>
            Проект ремонта
          </div>
          <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 28, color: COLORS.ink, margin: '4px 0 0' }}>
            Квартира Нурлы Тау
          </h1>
        </header>

        <nav style={{ display: 'flex', gap: 6, marginBottom: 20, flexWrap: 'wrap' }}>
          {tabs.map(({ id, label, icon: Icon }) => (
            <button
              key={id}
              onClick={() => setTab(id)}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: 6,
                padding: '8px 14px',
                borderRadius: 8,
                border: `1px solid ${tab === id ? COLORS.accent : COLORS.line}`,
                background: tab === id ? COLORS.accentSoft : COLORS.surface,
                color: tab === id ? COLORS.accent : COLORS.inkSoft,
                fontFamily: 'Inter, sans-serif',
                fontSize: 14,
                fontWeight: 500,
                cursor: 'pointer',
              }}
            >
              <Icon size={15} /> {label}
            </button>
          ))}
        </nav>

        {tab === 'overview' && (
          <Overview
            rooms={rooms}
            totalPlanned={totalPlanned}
            totalSpent={totalSpent}
            overallProgress={overallProgress}
            spentByRoom={spentByRoom}
            tasksByRoom={tasksByRoom}
            tasks={tasks}
            issues={issues}
            purchases={purchases}
          />
        )}
        {tab === 'budget' && (
          <Budget rooms={rooms} setRooms={setRooms} expenses={expenses} setExpenses={setExpenses} spentByRoom={spentByRoom} totalPlanned={totalPlanned} totalSpent={totalSpent} />
        )}
        {tab === 'plan' && <Plan rooms={rooms} tasks={tasks} setTasks={setTasks} />}
        {tab === 'purchases' && <Purchases rooms={rooms} purchases={purchases} setPurchases={setPurchases} setExpenses={setExpenses} />}
        {tab === 'issues' && <Issues rooms={rooms} issues={issues} setIssues={setIssues} />}
        {tab === 'documents' && <Documents documents={documents} setDocuments={setDocuments} />}
        {tab === 'advances' && <Advances advances={advances} setAdvances={setAdvances} />}
      </div>
    </div>
  );
}

function Overview({ rooms, totalPlanned, totalSpent, overallProgress, spentByRoom, tasksByRoom, tasks, issues, purchases }) {
  const stageStats = STAGES.map((stage) => {
    const stageTasks = tasks.filter((t) => t.stage === stage);
    const done = stageTasks.filter((t) => t.done).length;
    return { stage, total: stageTasks.length, done };
  });

  const roomName = (id) => (rooms.find((r) => r.id === id) || {}).name || '—';
  const inProgressIssues = (issues || []).filter((i) => i.status === 'В работе');

  const todayStr = new Date().toISOString().slice(0, 10);
  const in14Str = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000).toISOString().slice(0, 10);
  const upcomingPurchases = (purchases || [])
    .filter((p) => p.status !== 'Куплено' && p.deadline && p.deadline >= todayStr && p.deadline <= in14Str)
    .sort((a, b) => a.deadline.localeCompare(b.deadline));

  return (
    <div>
      <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', marginBottom: 20 }}>
        <StatCard label="Потрачено" value={fmt(totalSpent)} sub={`${totalPlanned ? Math.round((totalSpent / totalPlanned) * 100) : 0}% от бюджета`} />
        <StatCard label="Работы выполнены" value={`${overallProgress}%`} sub={`${tasks.filter((t) => t.done).length} из ${tasks.length} задач`} />
      </div>

      <BlueprintCard style={{ marginBottom: 20 }}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 15, fontWeight: 700, color: COLORS.ink, marginBottom: 14 }}>
          Прогресс по этапам
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          {stageStats.map(({ stage, total, done }) => (
            <div key={stage}>
              <div style={{ display: 'flex', justifyContent: 'space-between', fontFamily: 'Inter, sans-serif', fontSize: 13, color: COLORS.inkSoft, marginBottom: 4 }}>
                <span>{stage}</span>
                <span style={{ fontFamily: 'IBM Plex Mono, monospace' }}>{done}/{total || 0}</span>
              </div>
              <ProgressBar pct={total ? (done / total) * 100 : 0} color={COLORS.success} />
            </div>
          ))}
        </div>
      </BlueprintCard>

      <BlueprintCard style={{ marginBottom: 20 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 14 }}>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 15, fontWeight: 700, color: COLORS.ink }}>
            Вопросы в работе
          </div>
          <span style={{ fontFamily: 'IBM Plex Mono, monospace', fontSize: 12, color: COLORS.inkSoft }}>{inProgressIssues.length}</span>
        </div>
        {inProgressIssues.length === 0 ? (
          <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 13, color: COLORS.inkSoft, fontStyle: 'italic' }}>Нет вопросов со статусом «В работе»</div>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
            {inProgressIssues.map((i) => (
              <div key={i.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8, padding: '8px 0', borderBottom: `1px solid ${COLORS.bg}`, flexWrap: 'wrap' }}>
                <div>
                  <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 13, color: COLORS.ink, fontWeight: 500 }}>{i.name}</div>
                  <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 11, color: COLORS.inkSoft }}>{i.type}{i.roomId ? ` · ${roomName(i.roomId)}` : ''}</div>
                </div>
                <div style={{ fontFamily: 'IBM Plex Mono, monospace', fontSize: 12, color: COLORS.inkSoft }}>{formatDateRu(i.deadline)}</div>
              </div>
            ))}
          </div>
        )}
      </BlueprintCard>

      <BlueprintCard>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 14 }}>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 15, fontWeight: 700, color: COLORS.ink }}>
            Покупки с дедлайном через 1-2 недели
          </div>
          <span style={{ fontFamily: 'IBM Plex Mono, monospace', fontSize: 12, color: COLORS.inkSoft }}>{upcomingPurchases.length}</span>
        </div>
        {upcomingPurchases.length === 0 ? (
          <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 13, color: COLORS.inkSoft, fontStyle: 'italic' }}>Нет покупок с дедлайном в ближайшие две недели</div>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
            {upcomingPurchases.map((p) => (
              <div key={p.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8, padding: '8px 0', borderBottom: `1px solid ${COLORS.bg}`, flexWrap: 'wrap' }}>
                <div>
                  <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 13, color: COLORS.ink, fontWeight: 500 }}>{p.name}</div>
                  <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 11, color: COLORS.inkSoft }}>{p.type}{p.roomId ? ` · ${roomName(p.roomId)}` : ''} · {p.status}</div>
                </div>
                <div style={{ fontFamily: 'IBM Plex Mono, monospace', fontSize: 12, color: COLORS.inkSoft }}>{formatDateRu(p.deadline)}</div>
              </div>
            ))}
          </div>
        )}
      </BlueprintCard>
    </div>
  );
}

function Budget({ rooms, setRooms, expenses, setExpenses, spentByRoom, totalPlanned, totalSpent }) {
  const emptyForm = { name: '', type: PURCHASE_TYPES[0], subtype: EXPENSE_SUBTYPES[0], roomId: '', amount: '', receiptUrl: '', date: '', paymentComment: '' };
  const [form, setForm] = useState(emptyForm);
  const [collapsedGroups, setCollapsedGroups] = useState({});
  const [expandedIds, setExpandedIds] = useState({});
  const [editingIds, setEditingIds] = useState({});
  const [groupBy, setGroupBy] = useState('type');
  const [search, setSearch] = useState('');
  const [bulkOpen, setBulkOpen] = useState(false);
  const [bulkText, setBulkText] = useState('');
  const [bulkType, setBulkType] = useState(PURCHASE_TYPES[0]);
  const [bulkSubtype, setBulkSubtype] = useState(EXPENSE_SUBTYPES[1]);
  const [bulkRoomId, setBulkRoomId] = useState('');
  const [bulkDate, setBulkDate] = useState('');
  const [bulkComment, setBulkComment] = useState('');

  const bulkPreview = useMemo(() => {
    return bulkText.split('\n').map((line) => line.trim()).filter(Boolean).map((line) => {
      const parts = line.split('|').map((p) => p.trim());
      const name = parts[0] || '';
      const amount = Number((parts[1] || '').replace(/\s|₸/g, ''));
      return { name, amount: isNaN(amount) ? 0 : amount };
    }).filter((r) => r.name);
  }, [bulkText]);
  const bulkTotal = bulkPreview.reduce((s, r) => s + r.amount, 0);

  const addBulkExpenses = () => {
    if (bulkPreview.length === 0) return;
    const newOnes = bulkPreview.map((r) => ({
      id: uid('e'), name: r.name, type: bulkType, subtype: bulkSubtype, roomId: bulkRoomId,
      amount: r.amount, receiptUrl: '', date: bulkDate, paymentComment: bulkComment,
    }));
    setExpenses((prev) => [...prev, ...newOnes]);
    setBulkText('');
    setBulkOpen(false);
  };

  const roomName = (id) => (rooms.find((r) => r.id === id) || {}).name || '—';

  const worked = useMemo(() => expenses.filter((e) => e.subtype === 'Работа').reduce((s, e) => s + (Number(e.amount) || 0), 0), [expenses]);
  const materials = useMemo(() => expenses.filter((e) => e.subtype === 'Материалы').reduce((s, e) => s + (Number(e.amount) || 0), 0), [expenses]);

  const addExpense = () => {
    if (!form.name.trim() || !form.amount) return;
    setExpenses((prev) => [...prev, {
      id: uid('e'), name: form.name.trim(), type: form.type, subtype: form.subtype, roomId: form.roomId,
      amount: Number(form.amount), receiptUrl: form.receiptUrl.trim(), date: form.date, paymentComment: form.paymentComment.trim(),
    }]);
    setForm(emptyForm);
  };

  const updateField = (id, field, value) => {
    setExpenses((prev) => prev.map((e) => (e.id === id ? { ...e, [field]: value } : e)));
  };
  const removeExpense = (id) => setExpenses((prev) => prev.filter((e) => e.id !== id));
  const toggleExpand = (id) => setExpandedIds((prev) => ({ ...prev, [id]: !prev[id] }));
  const toggleEdit = (id) => setEditingIds((prev) => ({ ...prev, [id]: !prev[id] }));

  const updatePlanned = (id, value) => {
    setRooms((prev) => prev.map((r) => (r.id === id ? { ...r, planned: Number(value) || 0 } : r)));
  };

  const [editingRoomId, setEditingRoomId] = useState(null);
  const [editForm, setEditForm] = useState({});
  const [newRoomForm, setNewRoomForm] = useState({ name: '', area: '', planned: '' });
  const [confirmAction, setConfirmAction] = useState(null);
  const [roomsCollapsed, setRoomsCollapsed] = useState(true);

  const startEditRoom = (r) => { setEditingRoomId(r.id); setEditForm({ name: r.name, area: r.area }); };
  const saveEditRoom = (id) => {
    setRooms((prev) => prev.map((r) => (r.id === id ? { ...r, name: editForm.name.trim() || r.name, area: Number(editForm.area) || 0 } : r)));
    setEditingRoomId(null);
  };
  const removeRoom = (id) => {
    setRooms((prev) => prev.filter((r) => r.id !== id));
  };
  const addRoom = () => {
    if (!newRoomForm.name.trim()) return;
    setRooms((prev) => [...prev, { id: uid('room'), name: newRoomForm.name.trim(), area: Number(newRoomForm.area) || 0, planned: Number(newRoomForm.planned) || 0 }]);
    setNewRoomForm({ name: '', area: '', planned: '' });
  };

  const exportToExcel = () => {
    const roomsData = rooms.map((r) => ({
      'Комната': r.name,
      'Площадь, м²': r.area,
      'Бюджет, ₸': r.planned,
      'Потрачено, ₸': spentByRoom[r.id] || 0,
      'Остаток, ₸': (r.planned || 0) - (spentByRoom[r.id] || 0),
    }));
    roomsData.push({
      'Комната': 'ИТОГО',
      'Площадь, м²': rooms.reduce((s, r) => s + (Number(r.area) || 0), 0),
      'Бюджет, ₸': totalPlanned,
      'Потрачено, ₸': totalSpent,
      'Остаток, ₸': totalPlanned - totalSpent,
    });

    const expensesData = expenses.map((e) => ({
      'Название': e.name,
      'Тип': e.type,
      'Подтип': e.subtype,
      'Комната': roomName(e.roomId),
      'Сумма, ₸': e.amount,
      'Дата': formatDateRu(e.date),
      'Ссылка на чек': e.receiptUrl || '',
      'Комментарий по оплате': e.paymentComment || '',
    }));

    const wb = XLSX.utils.book_new();

    const ws2 = XLSX.utils.json_to_sheet(expensesData.length ? expensesData : [{ 'Название': '', 'Тип': '', 'Подтип': '', 'Комната': '', 'Сумма, ₸': '', 'Дата': '', 'Ссылка на чек': '', 'Комментарий по оплате': '' }]);
    ws2['!cols'] = [{ wch: 24 }, { wch: 20 }, { wch: 12 }, { wch: 18 }, { wch: 12 }, { wch: 16 }, { wch: 28 }, { wch: 24 }];
    XLSX.utils.book_append_sheet(wb, ws2, 'Расходы');

    const ws1 = XLSX.utils.json_to_sheet(roomsData);
    ws1['!cols'] = [{ wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, ws1, 'Смета по комнатам');

    const dateStr = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `Смета ремонта ${dateStr}.xlsx`);
  };

  const searchResults = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return null;
    const matches = expenses.filter((e) => e.name.toLowerCase().includes(q) || e.type.toLowerCase().includes(q) || e.subtype.toLowerCase().includes(q));
    const total = matches.reduce((s, e) => s + (Number(e.amount) || 0), 0);
    return { matches, total };
  }, [search, expenses]);

  const groups = useMemo(() => {
    if (groupBy === 'subtype') {
      return EXPENSE_SUBTYPES.map((s) => ({ key: s, label: s, items: expenses.filter((e) => e.subtype === s) }));
    }
    if (groupBy === 'room') {
      const roomGroups = rooms.map((r) => ({ key: r.id, label: r.name, items: expenses.filter((e) => e.roomId === r.id) }));
      const noRoom = expenses.filter((e) => !e.roomId);
      return [...roomGroups, { key: 'none', label: 'Без комнаты', items: noRoom }];
    }
    return PURCHASE_TYPES.map((t) => ({ key: t, label: t, items: expenses.filter((e) => e.type === t) }));
  }, [groupBy, expenses, rooms]);

  const thStyle = { textAlign: 'left', fontFamily: 'Inter, sans-serif', fontSize: 11, letterSpacing: '0.04em', textTransform: 'uppercase', color: COLORS.inkSoft, padding: '6px 8px', borderBottom: `1px solid ${COLORS.line}` };
  const tdStyle = { padding: '8px', verticalAlign: 'middle', borderBottom: `1px solid ${COLORS.bg}` };

  return (
    <div>
      <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 12 }}>
        <button
          onClick={exportToExcel}
          style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '8px 14px', borderRadius: 7, border: `1px solid ${COLORS.line}`, background: COLORS.surface, color: COLORS.ink, fontFamily: 'Inter, sans-serif', fontSize: 13, fontWeight: 500, cursor: 'pointer' }}
        >
          <Download size={14} /> Выгрузить в Excel
        </button>
      </div>

      <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', marginBottom: 16 }}>
        <StatCard label="Потрачено" value={fmt(totalSpent)} />
        <StatCard label="Работы" value={fmt(worked)} />
        <StatCard label="Материалы" value={fmt(materials)} />
      </div>

      <BlueprintCard style={{ marginBottom: 20 }}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 15, fontWeight: 700, color: COLORS.ink, marginBottom: 10 }}>
          Поиск по расходам
        </div>
        <TextInput placeholder="Например: техника, плитка, электрика…" value={search} onChange={(e) => setSearch(e.target.value)} />
        {searchResults && (
          <div style={{ marginTop: 12 }}>
            {searchResults.matches.length === 0 ? (
              <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 13, color: COLORS.inkSoft, fontStyle: 'italic' }}>
                Ничего не найдено по запросу «{search}»
              </div>
            ) : (
              <>
                <div style={{ fontFamily: 'IBM Plex Mono, monospace', fontSize: 14, color: COLORS.accent, fontWeight: 600, marginBottom: 8 }}>
                  Итого: {fmt(searchResults.total)} ({searchResults.matches.length} {searchResults.matches.length === 1 ? 'запись' : 'записей'})
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                  {searchResults.matches.map((e) => (
                    <div key={e.id} style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13, padding: '4px 0', borderBottom: `1px solid ${COLORS.bg}` }}>
                      <span style={{ color: COLORS.ink }}>{e.name} <span style={{ color: COLORS.inkSoft, fontSize: 11 }}>· {roomName(e.roomId)}</span></span>
                      <span style={{ fontFamily: 'IBM Plex Mono, monospace', color: COLORS.ink }}>{fmt(e.amount)}</span>
                    </div>
                  ))}
                </div>
              </>
            )}
          </div>
        )}
      </BlueprintCard>

      <BlueprintCard style={{ marginBottom: 20, padding: 0, overflow: 'hidden' }}>
        <div
          onClick={() => setBulkOpen(!bulkOpen)}
          style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer', padding: 16 }}
        >
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 15, fontWeight: 700, color: COLORS.ink }}>
            Массовый ввод расходов (из накладной/чека)
          </div>
          {bulkOpen ? <ChevronUp size={18} color={COLORS.inkSoft} /> : <ChevronDown size={18} color={COLORS.inkSoft} />}
        </div>
        {bulkOpen && (
          <div style={{ padding: '0 16px 16px' }}>
            <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 12, color: COLORS.inkSoft, marginBottom: 8 }}>
              По одной позиции в строке, формат: <b>Название | Сумма</b>. Общие Тип / Подтип / Комната / Дата / Комментарий применятся ко всем строкам сразу — проверь и поправь перед добавлением.
            </div>
            <TextArea maxLength={5000} rows={8} value={bulkText} onChange={(e) => setBulkText(e.target.value)} style={{ fontFamily: 'IBM Plex Mono, monospace', fontSize: 13 }} />
            <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginTop: 10 }}>
              <Select value={bulkType} onChange={(e) => setBulkType(e.target.value)} style={{ flex: '1 1 160px' }}>
                {PURCHASE_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
              </Select>
              <Select value={bulkSubtype} onChange={(e) => setBulkSubtype(e.target.value)} style={{ flex: '1 1 130px' }}>
                {EXPENSE_SUBTYPES.map((s) => <option key={s} value={s}>{s}</option>)}
              </Select>
              <Select value={bulkRoomId} onChange={(e) => setBulkRoomId(e.target.value)} style={{ flex: '1 1 140px' }}>
                <option value="">Без комнаты</option>
                {rooms.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
              </Select>
              <TextInput type="date" value={bulkDate} onChange={(e) => setBulkDate(e.target.value)} style={{ flex: '1 1 150px' }} />
              <TextInput placeholder="Комментарий по оплате (общий)" value={bulkComment} onChange={(e) => setBulkComment(e.target.value)} style={{ flex: '1 1 200px' }} />
            </div>
            <div style={{ fontFamily: 'IBM Plex Mono, monospace', fontSize: 13, color: COLORS.accent, fontWeight: 600, margin: '10px 0' }}>
              Распознано позиций: {bulkPreview.length} · Сумма: {fmt(bulkTotal)}
            </div>
            <button
              onClick={addBulkExpenses}
              disabled={bulkPreview.length === 0}
              style={{ display: 'flex', alignItems: 'center', gap: 4, padding: '8px 14px', borderRadius: 7, border: 'none', background: bulkPreview.length ? COLORS.accent : COLORS.line, color: '#fff', fontFamily: 'Inter, sans-serif', fontSize: 13, fontWeight: 500, cursor: bulkPreview.length ? 'pointer' : 'not-allowed' }}
            >
              <Plus size={14} /> Добавить все ({bulkPreview.length})
            </button>
          </div>
        )}
      </BlueprintCard>

      <BlueprintCard style={{ marginBottom: 20 }}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 15, fontWeight: 700, color: COLORS.ink, marginBottom: 12 }}>
          Новый расход
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 10 }}>
          <TextInput placeholder="Название" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} style={{ flex: '2 1 180px' }} />
          <Select value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })} style={{ flex: '1 1 160px' }}>
            {PURCHASE_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
          </Select>
          <Select value={form.subtype} onChange={(e) => setForm({ ...form, subtype: e.target.value })} style={{ flex: '1 1 130px' }}>
            {EXPENSE_SUBTYPES.map((s) => <option key={s} value={s}>{s}</option>)}
          </Select>
          <Select value={form.roomId} onChange={(e) => setForm({ ...form, roomId: e.target.value })} style={{ flex: '1 1 140px' }}>
            <option value="">Без комнаты</option>
            {rooms.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
          </Select>
          <TextInput type="number" placeholder="Сумма, ₸" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} style={{ flex: '1 1 110px' }} />
          <TextInput type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} style={{ flex: '1 1 150px' }} />
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 12 }}>
          <TextInput placeholder="Ссылка на чек" value={form.receiptUrl} onChange={(e) => setForm({ ...form, receiptUrl: e.target.value })} style={{ flex: '1 1 220px' }} />
          <TextInput placeholder="Комментарий по оплате" value={form.paymentComment} onChange={(e) => setForm({ ...form, paymentComment: e.target.value })} style={{ flex: '1 1 220px' }} />
        </div>
        <button
          onClick={addExpense}
          style={{ display: 'flex', alignItems: 'center', gap: 4, padding: '8px 14px', borderRadius: 7, border: 'none', background: COLORS.accent, color: '#fff', fontFamily: 'Inter, sans-serif', fontSize: 13, fontWeight: 500, cursor: 'pointer' }}
        >
          <Plus size={14} /> Добавить расход
        </button>
      </BlueprintCard>

      <div style={{ display: 'flex', gap: 6, marginBottom: 16, flexWrap: 'wrap', alignItems: 'center' }}>
        <span style={{ fontFamily: 'Inter, sans-serif', fontSize: 12, color: COLORS.inkSoft, marginRight: 2 }}>Группировать:</span>
        {[
          { id: 'type', label: 'По типу' },
          { id: 'subtype', label: 'По подтипу' },
          { id: 'room', label: 'По комнате' },
        ].map((g) => (
          <button
            key={g.id}
            onClick={() => setGroupBy(g.id)}
            style={{
              padding: '6px 12px', borderRadius: 999,
              border: `1px solid ${groupBy === g.id ? COLORS.accent : COLORS.line}`,
              background: groupBy === g.id ? COLORS.accentSoft : COLORS.surface,
              color: groupBy === g.id ? COLORS.accent : COLORS.inkSoft,
              fontFamily: 'Inter, sans-serif', fontSize: 12, fontWeight: 500, cursor: 'pointer',
            }}
          >
            {g.label}
          </button>
        ))}
      </div>

      {groups.map((group) => {
        const { key, label, items } = group;
        const groupSum = items.reduce((s, e) => s + (Number(e.amount) || 0), 0);
        const isCollapsed = collapsedGroups[key];
        return (
          <BlueprintCard key={key} style={{ marginBottom: 12, padding: 0, overflow: 'hidden' }}>
            <div
              onClick={() => setCollapsedGroups({ ...collapsedGroups, [key]: !isCollapsed })}
              style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer', padding: 16 }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <span style={{ display: 'inline-block', width: 10, height: 10, borderRadius: '50%', background: COLORS.accent }} />
                <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 15, fontWeight: 700, color: COLORS.ink }}>{label}</span>
                <span style={{ fontFamily: 'IBM Plex Mono, monospace', fontSize: 12, color: COLORS.inkSoft }}>{items.length} · {fmt(groupSum)}</span>
              </div>
              {isCollapsed ? <ChevronDown size={18} color={COLORS.inkSoft} /> : <ChevronUp size={18} color={COLORS.inkSoft} />}
            </div>

            {!isCollapsed && (
              items.length === 0 ? (
                <div style={{ padding: '0 16px 16px', fontFamily: 'Inter, sans-serif', fontSize: 13, color: COLORS.inkSoft, fontStyle: 'italic' }}>Пусто</div>
              ) : (
                <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                  <thead>
                    <tr>
                      <th style={{ ...thStyle, paddingLeft: 16 }}>Название</th>
                      <th style={thStyle}>Дата</th>
                      <th style={thStyle}>Сумма</th>
                      <th style={{ ...thStyle, width: 40, paddingRight: 16 }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {items.map((e) => {
                      const isOpen = !!expandedIds[e.id];
                      return (
                        <React.Fragment key={e.id}>
                          <tr onClick={() => toggleExpand(e.id)} style={{ cursor: 'pointer' }}>
                            <td style={{ ...tdStyle, paddingLeft: 16, fontFamily: 'Inter, sans-serif', fontSize: 13, color: COLORS.ink, fontWeight: 500 }}>{e.name}</td>
                            <td style={{ ...tdStyle, fontFamily: 'IBM Plex Mono, monospace', fontSize: 12, color: COLORS.inkSoft }}>{formatDateRu(e.date)}</td>
                            <td style={{ ...tdStyle, fontFamily: 'IBM Plex Mono, monospace', fontSize: 13, color: COLORS.ink }}>{fmt(e.amount)}</td>
                            <td style={{ ...tdStyle, paddingRight: 16, textAlign: 'right' }}>
                              {isOpen ? <ChevronUp size={16} color={COLORS.inkSoft} /> : <ChevronDown size={16} color={COLORS.inkSoft} />}
                            </td>
                          </tr>
                          {isOpen && (
                            <tr>
                              <td colSpan={4} style={{ padding: '0 16px 16px', borderBottom: `1px solid ${COLORS.bg}`, background: COLORS.bg }}>
                                <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', paddingTop: 12 }}>
                                  <div style={{ flex: '2 1 200px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Название</label>
                                    {editingIds[e.id] ? (
                                      <TextInput value={e.name} onChange={(ev) => updateField(e.id, 'name', ev.target.value)} />
                                    ) : (
                                      <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 14, color: COLORS.ink, padding: '8px 0' }}>{e.name}</div>
                                    )}
                                  </div>
                                  <div style={{ flex: '1 1 150px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Тип</label>
                                    <Select value={e.type} onChange={(ev) => updateField(e.id, 'type', ev.target.value)}>
                                      {PURCHASE_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
                                    </Select>
                                  </div>
                                  <div style={{ flex: '1 1 120px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Подтип</label>
                                    <Select value={e.subtype} onChange={(ev) => updateField(e.id, 'subtype', ev.target.value)}>
                                      {EXPENSE_SUBTYPES.map((s) => <option key={s} value={s}>{s}</option>)}
                                    </Select>
                                  </div>
                                  <div style={{ flex: '1 1 140px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Комната</label>
                                    <Select value={e.roomId} onChange={(ev) => updateField(e.id, 'roomId', ev.target.value)}>
                                      <option value="">Без комнаты</option>
                                      {rooms.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
                                    </Select>
                                  </div>
                                  <div style={{ flex: '1 1 110px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Сумма, ₸</label>
                                    <TextInput type="number" value={e.amount} onChange={(ev) => updateField(e.id, 'amount', ev.target.value)} />
                                  </div>
                                  <div style={{ flex: '1 1 150px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Дата</label>
                                    <TextInput type="date" value={e.date || ''} onChange={(ev) => updateField(e.id, 'date', ev.target.value)} />
                                  </div>
                                  <div style={{ flex: '1 1 220px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Ссылка на чек</label>
                                    <TextInput placeholder="https://…" value={e.receiptUrl || ''} onChange={(ev) => updateField(e.id, 'receiptUrl', ev.target.value)} />
                                    {e.receiptUrl && (
                                      <a href={e.receiptUrl} target="_blank" rel="noopener noreferrer" style={{ display: 'inline-flex', alignItems: 'center', gap: 4, color: COLORS.accent, fontSize: 11, marginTop: 4 }}>
                                        <Paperclip size={11} /> Открыть чек
                                      </a>
                                    )}
                                  </div>
                                  <div style={{ flex: '2 1 220px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Комментарий по оплате</label>
                                    <TextInput placeholder="Например: наличными, аванс 50%…" value={e.paymentComment || ''} onChange={(ev) => updateField(e.id, 'paymentComment', ev.target.value)} />
                                  </div>
                                </div>

                                <div style={{ marginTop: 12, display: 'flex', gap: 8 }}>
                                  <button
                                    onClick={() => toggleEdit(e.id)}
                                    style={{ display: 'flex', alignItems: 'center', gap: 4, padding: '6px 10px', borderRadius: 7, border: `1px solid ${editingIds[e.id] ? COLORS.accent : COLORS.line}`, background: editingIds[e.id] ? COLORS.accentSoft : COLORS.surface, color: editingIds[e.id] ? COLORS.accent : COLORS.ink, fontFamily: 'Inter, sans-serif', fontSize: 12, fontWeight: 500, cursor: 'pointer' }}
                                  >
                                    {editingIds[e.id] ? <><Check size={13} /> Готово</> : <><Pencil size={13} /> Редактировать</>}
                                  </button>
                                  <button
                                    onClick={() => removeExpense(e.id)}
                                    style={{ display: 'flex', alignItems: 'center', gap: 4, padding: '6px 10px', borderRadius: 7, border: `1px solid ${COLORS.line}`, background: COLORS.surface, color: '#B23A2C', fontFamily: 'Inter, sans-serif', fontSize: 12, cursor: 'pointer' }}
                                  >
                                    <Trash2 size={13} /> Удалить расход
                                  </button>
                                </div>
                              </td>
                            </tr>
                          )}
                        </React.Fragment>
                      );
                    })}
                  </tbody>
                </table>
              )
            )}
          </BlueprintCard>
        );
      })}

      <BlueprintCard style={{ marginTop: 16, padding: 0, overflow: 'hidden' }}>
        <div
          onClick={() => setRoomsCollapsed(!roomsCollapsed)}
          style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer', padding: 16 }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 14, fontWeight: 700, color: COLORS.ink }}>Комнаты</span>
            <span style={{ fontFamily: 'IBM Plex Mono, monospace', fontSize: 12, color: COLORS.inkSoft }}>{rooms.length}</span>
          </div>
          {roomsCollapsed ? <ChevronDown size={18} color={COLORS.inkSoft} /> : <ChevronUp size={18} color={COLORS.inkSoft} />}
        </div>
        {!roomsCollapsed && (
        <div style={{ padding: '0 16px 16px' }}>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: 14 }}>
          {rooms.map((r) => {
            const isEditing = editingRoomId === r.id;
            return isEditing ? (
              <div key={r.id} style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
                <TextInput placeholder="Название" value={editForm.name} onChange={(ev) => setEditForm({ ...editForm, name: ev.target.value })} style={{ flex: '2 1 140px' }} />
                <TextInput type="number" placeholder="Площадь, м²" value={editForm.area} onChange={(ev) => setEditForm({ ...editForm, area: ev.target.value })} style={{ flex: '1 1 100px' }} />
                <IconBtn onClick={() => saveEditRoom(r.id)} title="Сохранить"><Pencil size={14} /></IconBtn>
                <IconBtn onClick={() => setEditingRoomId(null)} title="Отмена"><X size={14} /></IconBtn>
              </div>
            ) : (
              <div key={r.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8, padding: '6px 0', borderBottom: `1px solid ${COLORS.bg}` }}>
                <div>
                  <span style={{ fontFamily: 'Inter, sans-serif', fontSize: 13, color: COLORS.ink }}>{r.name}</span>
                  <span style={{ fontFamily: 'IBM Plex Mono, monospace', fontSize: 11, color: COLORS.inkSoft, marginLeft: 8 }}>{r.area} м²</span>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Бюджет:</label>
                  <TextInput type="number" value={r.planned} onChange={(ev) => updatePlanned(r.id, ev.target.value)} style={{ width: 110 }} />
                  <IconBtn onClick={() => startEditRoom(r)} title="Редактировать"><Pencil size={14} /></IconBtn>
                  <IconBtn onClick={() => removeRoom(r.id)} title="Удалить комнату" danger><Trash2 size={14} /></IconBtn>
                </div>
              </div>
            );
          })}
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          <TextInput placeholder="Название" value={newRoomForm.name} onChange={(e) => setNewRoomForm({ ...newRoomForm, name: e.target.value })} style={{ flex: '2 1 140px' }} />
          <TextInput type="number" placeholder="Площадь, м²" value={newRoomForm.area} onChange={(e) => setNewRoomForm({ ...newRoomForm, area: e.target.value })} style={{ flex: '1 1 110px' }} />
          <TextInput type="number" placeholder="Бюджет, ₸" value={newRoomForm.planned} onChange={(e) => setNewRoomForm({ ...newRoomForm, planned: e.target.value })} style={{ flex: '1 1 110px' }} />
          <button
            onClick={addRoom}
            style={{ display: 'flex', alignItems: 'center', gap: 4, padding: '8px 14px', borderRadius: 7, border: 'none', background: COLORS.accent, color: '#fff', fontFamily: 'Inter, sans-serif', fontSize: 13, fontWeight: 500, cursor: 'pointer' }}
          >
            <Plus size={14} /> Добавить комнату
          </button>
        </div>
        </div>
        )}
      </BlueprintCard>

      <div style={{ marginTop: 10, display: 'flex', gap: 8, justifyContent: 'flex-end', flexWrap: 'wrap', alignItems: 'center' }}>
        {confirmAction ? (
          <>
            <span style={{ fontFamily: 'Inter, sans-serif', fontSize: 12, color: COLORS.inkSoft }}>
              {confirmAction === 'clear' ? 'Точно удалить все комнаты?' : 'Точно заменить на стандартный список (15 комнат)?'}
            </span>
            <button
              onClick={() => {
                if (confirmAction === 'clear') setRooms([]);
                else setRooms(DEFAULT_ROOMS);
                setConfirmAction(null);
              }}
              style={{ padding: '6px 12px', borderRadius: 7, border: 'none', background: '#B23A2C', color: '#fff', fontFamily: 'Inter, sans-serif', fontSize: 12, fontWeight: 500, cursor: 'pointer' }}
            >
              Да, подтвердить
            </button>
            <button
              onClick={() => setConfirmAction(null)}
              style={{ padding: '6px 12px', borderRadius: 7, border: `1px solid ${COLORS.line}`, background: 'transparent', color: COLORS.inkSoft, fontFamily: 'Inter, sans-serif', fontSize: 12, cursor: 'pointer' }}
            >
              Отмена
            </button>
          </>
        ) : (
          <>
            <button
              onClick={() => setConfirmAction('clear')}
              style={{ padding: '6px 12px', borderRadius: 7, border: `1px solid ${COLORS.line}`, background: 'transparent', color: '#B23A2C', fontFamily: 'Inter, sans-serif', fontSize: 12, cursor: 'pointer' }}
            >
              Удалить все комнаты
            </button>
            <button
              onClick={() => setConfirmAction('reset')}
              style={{ padding: '6px 12px', borderRadius: 7, border: `1px solid ${COLORS.line}`, background: 'transparent', color: COLORS.inkSoft, fontFamily: 'Inter, sans-serif', fontSize: 12, cursor: 'pointer' }}
            >
              Сбросить список комнат к стандартному
            </button>
          </>
        )}
      </div>
    </div>
  );
}

function Plan({ rooms, tasks, setTasks }) {
  const [form, setForm] = useState({ title: '', roomId: '', stage: STAGES[0], cost: '' });
  const [collapsed, setCollapsed] = useState({});

  const addTask = () => {
    if (!form.title.trim()) return;
    setTasks((prev) => [...prev, { id: uid('t'), title: form.title.trim(), roomId: form.roomId, stage: form.stage, cost: Number(form.cost) || 0, done: false }]);
    setForm({ ...form, title: '', cost: '' });
  };

  const toggleDone = (id) => {
    setTasks((prev) => prev.map((t) => (t.id === id ? { ...t, done: !t.done } : t)));
  };

  const removeTask = (id) => setTasks((prev) => prev.filter((t) => t.id !== id));

  const roomName = (id) => rooms.find((r) => r.id === id)?.name || '—';

  const totalDone = tasks.filter((t) => t.done).length;
  const overallPct = tasks.length ? Math.round((totalDone / tasks.length) * 100) : 0;

  return (
    <div>
      <BlueprintCard style={{ marginBottom: 20 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
          <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 11, letterSpacing: '0.06em', textTransform: 'uppercase', color: COLORS.inkSoft }}>
            % исполнения всех работ
          </div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 22, fontWeight: 700, color: COLORS.accent }}>{overallPct}%</div>
        </div>
        <ProgressBar pct={overallPct} color={COLORS.success} height={10} />
        <div style={{ fontFamily: 'IBM Plex Mono, monospace', fontSize: 12, color: COLORS.inkSoft, marginTop: 6 }}>{totalDone} из {tasks.length} пунктов выполнено</div>
      </BlueprintCard>

      <BlueprintCard style={{ marginBottom: 20 }}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 15, fontWeight: 700, color: COLORS.ink, marginBottom: 12 }}>
          Новый пункт плана
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          <TextInput placeholder="Что нужно сделать" value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })} style={{ flex: '2 1 180px' }} />
          <Select value={form.roomId} onChange={(e) => setForm({ ...form, roomId: e.target.value })} style={{ flex: '1 1 140px' }}>
            <option value="">Без комнаты</option>
            {rooms.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
          </Select>
          <Select value={form.stage} onChange={(e) => setForm({ ...form, stage: e.target.value })} style={{ flex: '1 1 160px' }}>
            {STAGES.map((s) => <option key={s} value={s}>{s}</option>)}
          </Select>
          <TextInput type="number" placeholder="Смета, ₸" value={form.cost} onChange={(e) => setForm({ ...form, cost: e.target.value })} style={{ flex: '1 1 100px' }} />
          <button
            onClick={addTask}
            style={{ display: 'flex', alignItems: 'center', gap: 4, padding: '8px 14px', borderRadius: 7, border: 'none', background: COLORS.accent, color: '#fff', fontFamily: 'Inter, sans-serif', fontSize: 13, fontWeight: 500, cursor: 'pointer' }}
          >
            <Plus size={14} /> Добавить
          </button>
        </div>
      </BlueprintCard>

      {STAGES.map((stage) => {
        const stageTasks = tasks.filter((t) => t.stage === stage);
        const done = stageTasks.filter((t) => t.done).length;
        const isCollapsed = collapsed[stage];
        const payment = STAGE_PAYMENT[stage];
        return (
          <BlueprintCard key={stage} style={{ marginBottom: 12 }}>
            <div
              onClick={() => setCollapsed({ ...collapsed, [stage]: !isCollapsed })}
              style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer' }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
                <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 15, fontWeight: 700, color: COLORS.ink }}>{stage}</span>
                {payment ? <span style={{ fontFamily: 'IBM Plex Mono, monospace', fontSize: 11, color: COLORS.inkSoft }}>· {fmt(payment)}</span> : null}
                <span style={{ fontFamily: 'IBM Plex Mono, monospace', fontSize: 12, color: COLORS.inkSoft }}>{done}/{stageTasks.length}</span>
              </div>
              {isCollapsed ? <ChevronDown size={18} color={COLORS.inkSoft} /> : <ChevronUp size={18} color={COLORS.inkSoft} />}
            </div>
            <div style={{ marginTop: 8 }}>
              <ProgressBar pct={stageTasks.length ? (done / stageTasks.length) * 100 : 0} color={COLORS.success} height={6} />
            </div>

            {!isCollapsed && (
              <div style={{ marginTop: 12, display: 'flex', flexDirection: 'column', gap: 2 }}>
                {stageTasks.length === 0 && (
                  <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 13, color: COLORS.inkSoft, fontStyle: 'italic' }}>Пунктов на этом этапе нет</div>
                )}
                {stageTasks.map((t) => (
                  <label key={t.id} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '8px 0', borderBottom: `1px solid ${COLORS.bg}`, cursor: 'pointer' }}>
                    <input
                      type="checkbox"
                      checked={!!t.done}
                      onChange={() => toggleDone(t.id)}
                      style={{ width: 18, height: 18, accentColor: COLORS.accent, cursor: 'pointer', flexShrink: 0 }}
                    />
                    <div style={{ flex: 1, minWidth: 140 }}>
                      <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 13, color: t.done ? COLORS.inkSoft : COLORS.ink, textDecoration: t.done ? 'line-through' : 'none' }}>
                        {t.title}
                      </div>
                      <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 11, color: COLORS.inkSoft }}>{roomName(t.roomId)} {t.cost ? `· ${fmt(t.cost)}` : ''}</div>
                    </div>
                    <IconBtn onClick={(e) => { e.preventDefault(); removeTask(t.id); }} title="Удалить" danger><Trash2 size={14} /></IconBtn>
                  </label>
                ))}
              </div>
            )}
          </BlueprintCard>
        );
      })}
    </div>
  );
}


function Purchases({ rooms, purchases, setPurchases, setExpenses }) {
  const emptyForm = { name: '', type: PURCHASE_TYPES[0], roomId: '', qty: '', price: '', status: PURCHASE_STATUSES[0], links: [''], deadline: '' };
  const [form, setForm] = useState(emptyForm);
  const [collapsedGroups, setCollapsedGroups] = useState({});
  const [expandedIds, setExpandedIds] = useState({});
  const [editingIds, setEditingIds] = useState({});
  const [groupBy, setGroupBy] = useState('status');

  const roomName = (id) => (rooms.find((r) => r.id === id) || {}).name || '—';

  const sumOf = (p) => {
    const qty = Number(p.qty);
    const price = Number(p.price);
    if (!qty || !price) return null;
    return qty * price;
  };

  const setFormLink = (index, value) => {
    setForm((f) => ({ ...f, links: f.links.map((l, i) => (i === index ? value : l)) }));
  };
  const addFormLink = () => setForm((f) => ({ ...f, links: [...f.links, ''] }));
  const removeFormLink = (index) => setForm((f) => ({ ...f, links: f.links.filter((_, i) => i !== index) }));

  const addPurchase = () => {
    if (!form.name.trim()) return;
    const links = form.links.map((l) => l.trim()).filter(Boolean);
    const newPurchase = { id: uid('p'), name: form.name.trim(), type: form.type, roomId: form.roomId, qty: form.qty, price: form.price, status: form.status, links, deadline: form.deadline };
    setPurchases((prev) => [...prev, newPurchase]);
    if (form.status === 'Куплено') createExpenseFromPurchase(newPurchase);
    setForm(emptyForm);
  };

  const createExpenseFromPurchase = (p) => {
    const qty = Number(p.qty);
    const price = Number(p.price);
    const amount = qty && price ? qty * price : 0;
    const today = new Date().toISOString().slice(0, 10);
    setExpenses((prevExp) => [...prevExp, {
      id: uid('e'),
      name: p.name,
      type: p.type,
      subtype: 'Материалы',
      roomId: p.roomId,
      amount,
      receiptUrl: '',
      date: today,
      paymentComment: `Создано автоматически из покупки «${p.name}»`,
    }]);
  };

  const updateField = (id, field, value) => {
    if (field === 'status' && value === 'Куплено') {
      const p = purchases.find((x) => x.id === id);
      if (p && p.status !== 'Куплено') createExpenseFromPurchase(p);
    }
    setPurchases((prev) => prev.map((p) => (p.id === id ? { ...p, [field]: value } : p)));
  };

  const updateLink = (id, index, value) => {
    setPurchases((prev) => prev.map((p) => (p.id === id ? { ...p, links: (p.links || []).map((l, i) => (i === index ? value : l)) } : p)));
  };
  const addLink = (id) => {
    setPurchases((prev) => prev.map((p) => (p.id === id ? { ...p, links: [...(p.links || []), ''] } : p)));
  };
  const removeLink = (id, index) => {
    setPurchases((prev) => prev.map((p) => (p.id === id ? { ...p, links: (p.links || []).filter((_, i) => i !== index) } : p)));
  };

  const removePurchase = (id) => setPurchases((prev) => prev.filter((p) => p.id !== id));
  const toggleExpand = (id) => setExpandedIds((prev) => ({ ...prev, [id]: !prev[id] }));
  const toggleEdit = (id) => setEditingIds((prev) => ({ ...prev, [id]: !prev[id] }));

  const groups = useMemo(() => {
    if (groupBy === 'type') {
      return PURCHASE_TYPES.map((t) => ({ key: t, label: t, items: purchases.filter((p) => p.type === t) }));
    }
    if (groupBy === 'room') {
      const roomGroups = rooms.map((r) => ({ key: r.id, label: r.name, items: purchases.filter((p) => p.roomId === r.id) }));
      const noRoom = purchases.filter((p) => !p.roomId);
      return [...roomGroups, { key: 'none', label: 'Без комнаты', items: noRoom }];
    }
    return PURCHASE_STATUSES.map((s) => ({ key: s, label: s, items: purchases.filter((p) => p.status === s), color: PURCHASE_STATUS_META[s].color }));
  }, [groupBy, purchases, rooms]);

  const bought = purchases.filter((p) => p.status === 'Куплено').reduce((s, p) => s + (sumOf(p) || 0), 0);
  const remaining = purchases.filter((p) => p.status !== 'Куплено').reduce((s, p) => s + (sumOf(p) || 0), 0);
  const totalAll = bought + remaining;
  const percentBought = totalAll ? Math.round((bought / totalAll) * 100) : 0;

  const thStyle = { textAlign: 'left', fontFamily: 'Inter, sans-serif', fontSize: 11, letterSpacing: '0.04em', textTransform: 'uppercase', color: COLORS.inkSoft, padding: '6px 8px', borderBottom: `1px solid ${COLORS.line}` };
  const tdStyle = { padding: '8px', verticalAlign: 'middle', borderBottom: `1px solid ${COLORS.bg}` };

  return (
    <div>
      <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', marginBottom: 20 }}>
        <StatCard label="Куплено" value={fmt(bought)} />
        <StatCard label="Остаток" value={fmt(remaining)} />
        <StatCard label="% закупленного" value={`${percentBought}%`} />
      </div>

      <div style={{ display: 'flex', gap: 6, marginBottom: 20, flexWrap: 'wrap', alignItems: 'center' }}>
        <span style={{ fontFamily: 'Inter, sans-serif', fontSize: 12, color: COLORS.inkSoft, marginRight: 2 }}>Группировать:</span>
        {[
          { id: 'status', label: 'По статусу' },
          { id: 'type', label: 'По типу' },
          { id: 'room', label: 'По комнате' },
        ].map((g) => (
          <button
            key={g.id}
            onClick={() => setGroupBy(g.id)}
            style={{
              padding: '6px 12px',
              borderRadius: 999,
              border: `1px solid ${groupBy === g.id ? COLORS.accent : COLORS.line}`,
              background: groupBy === g.id ? COLORS.accentSoft : COLORS.surface,
              color: groupBy === g.id ? COLORS.accent : COLORS.inkSoft,
              fontFamily: 'Inter, sans-serif', fontSize: 12, fontWeight: 500, cursor: 'pointer',
            }}
          >
            {g.label}
          </button>
        ))}
      </div>

      <BlueprintCard style={{ marginBottom: 20 }}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 15, fontWeight: 700, color: COLORS.ink, marginBottom: 12 }}>
          Новая покупка
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 10 }}>
          <TextInput placeholder="Название покупки" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} style={{ flex: '2 1 180px' }} />
          <Select value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })} style={{ flex: '1 1 160px' }}>
            {PURCHASE_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
          </Select>
          <Select value={form.roomId} onChange={(e) => setForm({ ...form, roomId: e.target.value })} style={{ flex: '1 1 140px' }}>
            <option value="">Без комнаты</option>
            {rooms.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
          </Select>
          <TextInput type="number" placeholder="Кол-во" value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} style={{ flex: '1 1 90px' }} />
          <TextInput type="number" placeholder="Цена, ₸" value={form.price} onChange={(e) => setForm({ ...form, price: e.target.value })} style={{ flex: '1 1 100px' }} />
          <Select value={form.status} onChange={(e) => setForm({ ...form, status: e.target.value })} style={{ flex: '1 1 150px' }}>
            {PURCHASE_STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
          </Select>
          <TextInput type="date" value={form.deadline} onChange={(e) => setForm({ ...form, deadline: e.target.value })} style={{ flex: '1 1 150px' }} />
        </div>

        <div style={{ marginBottom: 12 }}>
          <label style={{ fontFamily: 'Inter, sans-serif', fontSize: 12, color: COLORS.inkSoft }}>Ссылки</label>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginTop: 6 }}>
            {form.links.map((link, i) => (
              <div key={i} style={{ display: 'flex', gap: 6 }}>
                <TextInput placeholder="https://…" value={link} onChange={(e) => setFormLink(i, e.target.value)} />
                {form.links.length > 1 && (
                  <IconBtn onClick={() => removeFormLink(i)} title="Убрать ссылку" danger><X size={14} /></IconBtn>
                )}
              </div>
            ))}
            <button
              onClick={addFormLink}
              style={{ alignSelf: 'flex-start', display: 'flex', alignItems: 'center', gap: 4, padding: '5px 10px', borderRadius: 7, border: `1px dashed ${COLORS.line}`, background: 'transparent', color: COLORS.inkSoft, fontFamily: 'Inter, sans-serif', fontSize: 12, cursor: 'pointer' }}
            >
              <Plus size={12} /> Добавить ссылку
            </button>
          </div>
        </div>

        <button
          onClick={addPurchase}
          style={{ display: 'flex', alignItems: 'center', gap: 4, padding: '8px 14px', borderRadius: 7, border: 'none', background: COLORS.accent, color: '#fff', fontFamily: 'Inter, sans-serif', fontSize: 13, fontWeight: 500, cursor: 'pointer' }}
        >
          <Plus size={14} /> Добавить покупку
        </button>
      </BlueprintCard>

      {groups.map((group) => {
        const { key, label, items } = group;
        const groupSum = items.reduce((s, p) => s + (sumOf(p) || 0), 0);
        const isCollapsed = collapsedGroups[key];
        const dotColor = group.color || COLORS.accent;
        return (
          <BlueprintCard key={key} style={{ marginBottom: 12, padding: 0, overflow: 'hidden' }}>
            <div
              onClick={() => setCollapsedGroups({ ...collapsedGroups, [key]: !isCollapsed })}
              style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer', padding: 16 }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <span style={{ display: 'inline-block', width: 10, height: 10, borderRadius: '50%', background: dotColor }} />
                <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 15, fontWeight: 700, color: COLORS.ink }}>{label}</span>
                <span style={{ fontFamily: 'IBM Plex Mono, monospace', fontSize: 12, color: COLORS.inkSoft }}>{items.length} · {fmt(groupSum)}</span>
              </div>
              {isCollapsed ? <ChevronDown size={18} color={COLORS.inkSoft} /> : <ChevronUp size={18} color={COLORS.inkSoft} />}
            </div>

            {!isCollapsed && (
              items.length === 0 ? (
                <div style={{ padding: '0 16px 16px', fontFamily: 'Inter, sans-serif', fontSize: 13, color: COLORS.inkSoft, fontStyle: 'italic' }}>Пусто</div>
              ) : (
                <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                  <thead>
                    <tr>
                      <th style={{ ...thStyle, paddingLeft: 16 }}>Название</th>
                      <th style={thStyle}>Комната</th>
                      <th style={{ ...thStyle, width: 40, paddingRight: 16 }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {items.map((p) => {
                      const sum = sumOf(p);
                      const isOpen = !!expandedIds[p.id];
                      const links = p.links || [];
                      return (
                        <React.Fragment key={p.id}>
                          <tr onClick={() => toggleExpand(p.id)} style={{ cursor: 'pointer' }}>
                            <td style={{ ...tdStyle, paddingLeft: 16, fontFamily: 'Inter, sans-serif', fontSize: 13, color: COLORS.ink, fontWeight: 500 }}>{p.name}</td>
                            <td style={{ ...tdStyle, fontFamily: 'Inter, sans-serif', fontSize: 13, color: COLORS.inkSoft }}>{roomName(p.roomId)}</td>
                            <td style={{ ...tdStyle, paddingRight: 16, textAlign: 'right' }}>
                              {isOpen ? <ChevronUp size={16} color={COLORS.inkSoft} /> : <ChevronDown size={16} color={COLORS.inkSoft} />}
                            </td>
                          </tr>
                          {isOpen && (
                            <tr>
                              <td colSpan={3} style={{ padding: '0 16px 16px', borderBottom: `1px solid ${COLORS.bg}`, background: COLORS.bg }}>
                                <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', paddingTop: 12 }}>
                                  <div style={{ flex: '2 1 200px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Название</label>
                                    {editingIds[p.id] ? (
                                      <TextInput value={p.name} onChange={(e) => updateField(p.id, 'name', e.target.value)} />
                                    ) : (
                                      <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 14, color: COLORS.ink, padding: '8px 0' }}>{p.name}</div>
                                    )}
                                  </div>
                                  <div style={{ flex: '1 1 150px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Тип</label>
                                    <Select value={p.type} onChange={(e) => updateField(p.id, 'type', e.target.value)}>
                                      {PURCHASE_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
                                    </Select>
                                  </div>
                                  <div style={{ flex: '1 1 140px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Комната</label>
                                    <Select value={p.roomId} onChange={(e) => updateField(p.id, 'roomId', e.target.value)}>
                                      <option value="">Без комнаты</option>
                                      {rooms.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
                                    </Select>
                                  </div>
                                  <div style={{ flex: '1 1 80px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Кол-во</label>
                                    <TextInput type="number" value={p.qty} onChange={(e) => updateField(p.id, 'qty', e.target.value)} />
                                  </div>
                                  <div style={{ flex: '1 1 100px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Цена, ₸</label>
                                    <TextInput type="number" value={p.price} onChange={(e) => updateField(p.id, 'price', e.target.value)} />
                                  </div>
                                  <div style={{ flex: '1 1 100px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Сумма</label>
                                    <div style={{ fontFamily: 'IBM Plex Mono, monospace', fontSize: 14, color: COLORS.ink, padding: '8px 0' }}>{sum === null ? '—' : fmt(sum)}</div>
                                  </div>
                                  <div style={{ flex: '1 1 150px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Статус</label>
                                    <Select value={p.status} onChange={(e) => updateField(p.id, 'status', e.target.value)}>
                                      {PURCHASE_STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
                                    </Select>
                                  </div>
                                  <div style={{ flex: '1 1 150px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Дедлайн</label>
                                    <TextInput type="date" value={p.deadline || ''} onChange={(e) => updateField(p.id, 'deadline', e.target.value)} />
                                  </div>
                                </div>

                                <div style={{ marginTop: 10 }}>
                                  <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Ссылки</label>
                                  <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginTop: 6 }}>
                                    {links.length === 0 && (
                                      <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 12, color: COLORS.inkSoft, fontStyle: 'italic' }}>Ссылок нет</div>
                                    )}
                                    {links.map((link, i) => (
                                      <div key={i} style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                                        {link ? (
                                          <a href={link} target="_blank" rel="noopener noreferrer" style={{ display: 'flex', alignItems: 'center', gap: 4, color: COLORS.accent, fontSize: 12, fontFamily: 'Inter, sans-serif', flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                            <Link2 size={12} /> {link}
                                          </a>
                                        ) : (
                                          <TextInput placeholder="https://…" value={link} onChange={(e) => updateLink(p.id, i, e.target.value)} />
                                        )}
                                        <IconBtn onClick={() => removeLink(p.id, i)} title="Удалить ссылку" danger><X size={14} /></IconBtn>
                                      </div>
                                    ))}
                                    <button
                                      onClick={() => addLink(p.id)}
                                      style={{ alignSelf: 'flex-start', display: 'flex', alignItems: 'center', gap: 4, padding: '5px 10px', borderRadius: 7, border: `1px dashed ${COLORS.line}`, background: 'transparent', color: COLORS.inkSoft, fontFamily: 'Inter, sans-serif', fontSize: 12, cursor: 'pointer' }}
                                    >
                                      <Plus size={12} /> Добавить ссылку
                                    </button>
                                  </div>
                                </div>

                                <div style={{ marginTop: 12, display: 'flex', gap: 8 }}>
                                  <button
                                    onClick={() => toggleEdit(p.id)}
                                    style={{ display: 'flex', alignItems: 'center', gap: 4, padding: '6px 10px', borderRadius: 7, border: `1px solid ${editingIds[p.id] ? COLORS.accent : COLORS.line}`, background: editingIds[p.id] ? COLORS.accentSoft : COLORS.surface, color: editingIds[p.id] ? COLORS.accent : COLORS.ink, fontFamily: 'Inter, sans-serif', fontSize: 12, fontWeight: 500, cursor: 'pointer' }}
                                  >
                                    {editingIds[p.id] ? <><Check size={13} /> Готово</> : <><Pencil size={13} /> Редактировать</>}
                                  </button>
                                  <button
                                    onClick={() => removePurchase(p.id)}
                                    style={{ display: 'flex', alignItems: 'center', gap: 4, padding: '6px 10px', borderRadius: 7, border: `1px solid ${COLORS.line}`, background: COLORS.surface, color: '#B23A2C', fontFamily: 'Inter, sans-serif', fontSize: 12, cursor: 'pointer' }}
                                  >
                                    <Trash2 size={13} /> Удалить покупку
                                  </button>
                                </div>
                              </td>
                            </tr>
                          )}
                        </React.Fragment>
                      );
                    })}
                  </tbody>
                </table>
              )
            )}
          </BlueprintCard>
        );
      })}
    </div>
  );
}

function Issues({ rooms, issues, setIssues }) {
  const emptyForm = { name: '', type: PURCHASE_TYPES[0], roomId: rooms[0]?.id || '', comment: '', deadline: '', solution: '', status: ISSUE_STATUSES[0], responsible: RESPONSIBLE_PEOPLE[0] };
  const [form, setForm] = useState(emptyForm);
  const [collapsedGroups, setCollapsedGroups] = useState({});
  const [expandedIds, setExpandedIds] = useState({});
  const [editingIds, setEditingIds] = useState({});
  const [groupBy, setGroupBy] = useState('status');

  const roomName = (id) => (rooms.find((r) => r.id === id) || {}).name || '—';

  const addIssue = () => {
    if (!form.name.trim()) return;
    setIssues((prev) => [...prev, {
      id: uid('i'), name: form.name.trim(), type: form.type, roomId: form.roomId,
      comment: form.comment, deadline: form.deadline, solution: form.solution, status: form.status, responsible: form.responsible,
    }]);
    setForm(emptyForm);
  };

  const updateField = (id, field, value) => {
    setIssues((prev) => prev.map((i) => (i.id === id ? { ...i, [field]: value } : i)));
  };
  const removeIssue = (id) => setIssues((prev) => prev.filter((i) => i.id !== id));
  const toggleExpand = (id) => setExpandedIds((prev) => ({ ...prev, [id]: !prev[id] }));
  const toggleEdit = (id) => setEditingIds((prev) => ({ ...prev, [id]: !prev[id] }));

  const groups = useMemo(() => {
    if (groupBy === 'type') {
      return PURCHASE_TYPES.map((t) => ({ key: t, label: t, items: issues.filter((i) => i.type === t) }));
    }
    if (groupBy === 'status') {
      return ISSUE_STATUSES.map((s) => ({ key: s, label: s, items: issues.filter((i) => i.status === s), color: ISSUE_STATUS_META[s].color }));
    }
    const roomGroups = rooms.map((r) => ({ key: r.id, label: r.name, items: issues.filter((i) => i.roomId === r.id) }));
    const noRoom = issues.filter((i) => !i.roomId);
    return [...roomGroups, { key: 'none', label: 'Без комнаты', items: noRoom }];
  }, [groupBy, issues, rooms]);

  const thStyle = { textAlign: 'left', fontFamily: 'Inter, sans-serif', fontSize: 11, letterSpacing: '0.04em', textTransform: 'uppercase', color: COLORS.inkSoft, padding: '6px 8px', borderBottom: `1px solid ${COLORS.line}` };
  const tdStyle = { padding: '8px', verticalAlign: 'middle', borderBottom: `1px solid ${COLORS.bg}` };

  return (
    <div>
      <BlueprintCard style={{ marginBottom: 20 }}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 15, fontWeight: 700, color: COLORS.ink, marginBottom: 12 }}>
          Новый вопрос
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 10 }}>
          <TextInput placeholder="Название" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} style={{ flex: '2 1 180px' }} />
          <Select value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })} style={{ flex: '1 1 160px' }}>
            {PURCHASE_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
          </Select>
          <Select value={form.roomId} onChange={(e) => setForm({ ...form, roomId: e.target.value })} style={{ flex: '1 1 140px' }}>
            <option value="">Без комнаты</option>
            {rooms.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
          </Select>
          <TextInput type="date" value={form.deadline} onChange={(e) => setForm({ ...form, deadline: e.target.value })} style={{ flex: '1 1 150px' }} />
          <Select value={form.status} onChange={(e) => setForm({ ...form, status: e.target.value })} style={{ flex: '1 1 140px' }}>
            {ISSUE_STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
          </Select>
          <Select value={form.responsible} onChange={(e) => setForm({ ...form, responsible: e.target.value })} style={{ flex: '1 1 130px' }}>
            {RESPONSIBLE_PEOPLE.map((p) => <option key={p} value={p}>{p}</option>)}
          </Select>
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 12 }}>
          <div style={{ flex: '1 1 220px' }}>
            <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Комментарий</label>
            <TextArea maxLength={2000} rows={2} value={form.comment} onChange={(e) => setForm({ ...form, comment: e.target.value })} />
          </div>
          <div style={{ flex: '1 1 220px' }}>
            <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Решение</label>
            <TextArea maxLength={2000} rows={2} value={form.solution} onChange={(e) => setForm({ ...form, solution: e.target.value })} />
          </div>
        </div>
        <button
          onClick={addIssue}
          style={{ display: 'flex', alignItems: 'center', gap: 4, padding: '8px 14px', borderRadius: 7, border: 'none', background: COLORS.accent, color: '#fff', fontFamily: 'Inter, sans-serif', fontSize: 13, fontWeight: 500, cursor: 'pointer' }}
        >
          <Plus size={14} /> Добавить вопрос
        </button>
      </BlueprintCard>

      <div style={{ display: 'flex', gap: 6, marginBottom: 16, flexWrap: 'wrap', alignItems: 'center' }}>
        <span style={{ fontFamily: 'Inter, sans-serif', fontSize: 12, color: COLORS.inkSoft, marginRight: 2 }}>Группировать:</span>
        {[
          { id: 'status', label: 'По статусу' },
          { id: 'type', label: 'По типу' },
          { id: 'room', label: 'По комнате' },
        ].map((g) => (
          <button
            key={g.id}
            onClick={() => setGroupBy(g.id)}
            style={{
              padding: '6px 12px', borderRadius: 999,
              border: `1px solid ${groupBy === g.id ? COLORS.accent : COLORS.line}`,
              background: groupBy === g.id ? COLORS.accentSoft : COLORS.surface,
              color: groupBy === g.id ? COLORS.accent : COLORS.inkSoft,
              fontFamily: 'Inter, sans-serif', fontSize: 12, fontWeight: 500, cursor: 'pointer',
            }}
          >
            {g.label}
          </button>
        ))}
      </div>

      {groups.map((group) => {
        const { key, label, items } = group;
        const isCollapsed = collapsedGroups[key];
        const dotColor = group.color || COLORS.accent;
        return (
          <BlueprintCard key={key} style={{ marginBottom: 12, padding: 0, overflow: 'hidden' }}>
            <div
              onClick={() => setCollapsedGroups({ ...collapsedGroups, [key]: !isCollapsed })}
              style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer', padding: 16 }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <span style={{ display: 'inline-block', width: 10, height: 10, borderRadius: '50%', background: dotColor }} />
                <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 15, fontWeight: 700, color: COLORS.ink }}>{label}</span>
                <span style={{ fontFamily: 'IBM Plex Mono, monospace', fontSize: 12, color: COLORS.inkSoft }}>{items.length}</span>
              </div>
              {isCollapsed ? <ChevronDown size={18} color={COLORS.inkSoft} /> : <ChevronUp size={18} color={COLORS.inkSoft} />}
            </div>

            {!isCollapsed && (
              items.length === 0 ? (
                <div style={{ padding: '0 16px 16px', fontFamily: 'Inter, sans-serif', fontSize: 13, color: COLORS.inkSoft, fontStyle: 'italic' }}>Пусто</div>
              ) : (
                <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                  <thead>
                    <tr>
                      <th style={{ ...thStyle, paddingLeft: 16 }}>Название</th>
                      <th style={thStyle}>Дедлайн</th>
                      <th style={{ ...thStyle, width: 40, paddingRight: 16 }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {items.map((i) => {
                      const isOpen = !!expandedIds[i.id];
                      return (
                        <React.Fragment key={i.id}>
                          <tr onClick={() => toggleExpand(i.id)} style={{ cursor: 'pointer' }}>
                            <td style={{ ...tdStyle, paddingLeft: 16, fontFamily: 'Inter, sans-serif', fontSize: 13, color: COLORS.ink, fontWeight: 500 }}>{i.name}</td>
                            <td style={{ ...tdStyle, fontFamily: 'IBM Plex Mono, monospace', fontSize: 12, color: COLORS.inkSoft }}>{formatDateRu(i.deadline)}</td>
                            <td style={{ ...tdStyle, paddingRight: 16, textAlign: 'right' }}>
                              {isOpen ? <ChevronUp size={16} color={COLORS.inkSoft} /> : <ChevronDown size={16} color={COLORS.inkSoft} />}
                            </td>
                          </tr>
                          {isOpen && (
                            <tr>
                              <td colSpan={3} style={{ padding: '0 16px 16px', borderBottom: `1px solid ${COLORS.bg}`, background: COLORS.bg }}>
                                <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', paddingTop: 12 }}>
                                  <div style={{ flex: '2 1 200px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Название</label>
                                    {editingIds[i.id] ? (
                                      <TextInput value={i.name} onChange={(ev) => updateField(i.id, 'name', ev.target.value)} />
                                    ) : (
                                      <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 14, color: COLORS.ink, padding: '8px 0' }}>{i.name}</div>
                                    )}
                                  </div>
                                  <div style={{ flex: '1 1 150px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Тип</label>
                                    <Select value={i.type} onChange={(ev) => updateField(i.id, 'type', ev.target.value)}>
                                      {PURCHASE_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
                                    </Select>
                                  </div>
                                  <div style={{ flex: '1 1 140px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Комната</label>
                                    <Select value={i.roomId} onChange={(ev) => updateField(i.id, 'roomId', ev.target.value)}>
                                      <option value="">Без комнаты</option>
                                      {rooms.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
                                    </Select>
                                  </div>
                                  <div style={{ flex: '1 1 150px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Дедлайн</label>
                                    <TextInput type="date" value={i.deadline || ''} onChange={(ev) => updateField(i.id, 'deadline', ev.target.value)} />
                                  </div>
                                  <div style={{ flex: '1 1 140px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Статус</label>
                                    <Select value={i.status || ISSUE_STATUSES[0]} onChange={(ev) => updateField(i.id, 'status', ev.target.value)}>
                                      {ISSUE_STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
                                    </Select>
                                  </div>
                                  <div style={{ flex: '1 1 130px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Ответственный</label>
                                    <Select value={i.responsible || RESPONSIBLE_PEOPLE[0]} onChange={(ev) => updateField(i.id, 'responsible', ev.target.value)}>
                                      {RESPONSIBLE_PEOPLE.map((p) => <option key={p} value={p}>{p}</option>)}
                                    </Select>
                                  </div>
                                  <div style={{ flex: '1 1 220px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Комментарий</label>
                                    <TextArea maxLength={2000} rows={2} value={i.comment || ''} onChange={(ev) => updateField(i.id, 'comment', ev.target.value)} />
                                  </div>
                                  <div style={{ flex: '1 1 220px' }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Решение</label>
                                    <TextArea maxLength={2000} rows={2} value={i.solution || ''} onChange={(ev) => updateField(i.id, 'solution', ev.target.value)} />
                                  </div>
                                </div>

                                <div style={{ marginTop: 12, display: 'flex', gap: 8 }}>
                                  <button
                                    onClick={() => toggleEdit(i.id)}
                                    style={{ display: 'flex', alignItems: 'center', gap: 4, padding: '6px 10px', borderRadius: 7, border: `1px solid ${editingIds[i.id] ? COLORS.accent : COLORS.line}`, background: editingIds[i.id] ? COLORS.accentSoft : COLORS.surface, color: editingIds[i.id] ? COLORS.accent : COLORS.ink, fontFamily: 'Inter, sans-serif', fontSize: 12, fontWeight: 500, cursor: 'pointer' }}
                                  >
                                    {editingIds[i.id] ? <><Check size={13} /> Готово</> : <><Pencil size={13} /> Редактировать</>}
                                  </button>
                                  <button
                                    onClick={() => removeIssue(i.id)}
                                    style={{ display: 'flex', alignItems: 'center', gap: 4, padding: '6px 10px', borderRadius: 7, border: `1px solid ${COLORS.line}`, background: COLORS.surface, color: '#B23A2C', fontFamily: 'Inter, sans-serif', fontSize: 12, cursor: 'pointer' }}
                                  >
                                    <Trash2 size={13} /> Удалить вопрос
                                  </button>
                                </div>
                              </td>
                            </tr>
                          )}
                        </React.Fragment>
                      );
                    })}
                  </tbody>
                </table>
              )
            )}
          </BlueprintCard>
        );
      })}
    </div>
  );
}

function Advances({ advances, setAdvances }) {
  const emptyForm = { date: '', amount: '', status: ADVANCE_STATUSES[0], comment: '' };
  const [form, setForm] = useState(emptyForm);
  const [collapsedGroups, setCollapsedGroups] = useState({});
  const [expandedIds, setExpandedIds] = useState({});
  const [editingIds, setEditingIds] = useState({});

  const addAdvance = () => {
    if (!form.date || !form.amount) return;
    setAdvances((prev) => [...prev, { id: uid('a'), date: form.date, amount: Number(form.amount), status: form.status, comment: form.comment }]);
    setForm(emptyForm);
  };

  const updateField = (id, field, value) => {
    setAdvances((prev) => prev.map((a) => (a.id === id ? { ...a, [field]: value } : a)));
  };
  const removeAdvance = (id) => setAdvances((prev) => prev.filter((a) => a.id !== id));
  const toggleExpand = (id) => setExpandedIds((prev) => ({ ...prev, [id]: !prev[id] }));
  const toggleEdit = (id) => setEditingIds((prev) => ({ ...prev, [id]: !prev[id] }));

  const groups = useMemo(() => {
    return ADVANCE_STATUSES.map((s) => ({
      key: s, label: s, items: advances.filter((a) => a.status === s), color: ADVANCE_STATUS_META[s].color,
      sum: advances.filter((a) => a.status === s).reduce((sum, a) => sum + (Number(a.amount) || 0), 0),
    }));
  }, [advances]);

  const thStyle = { textAlign: 'left', fontFamily: 'Inter, sans-serif', fontSize: 11, letterSpacing: '0.04em', textTransform: 'uppercase', color: COLORS.inkSoft, padding: '6px 8px', borderBottom: `1px solid ${COLORS.line}` };
  const tdStyle = { padding: '8px', verticalAlign: 'middle', borderBottom: `1px solid ${COLORS.bg}` };

  return (
    <div>
      <BlueprintCard style={{ marginBottom: 20 }}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 15, fontWeight: 700, color: COLORS.ink, marginBottom: 12 }}>
          Новый аванс
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 10 }}>
          <TextInput type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} style={{ flex: '1 1 150px' }} />
          <TextInput type="number" placeholder="Сумма, ₸" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} style={{ flex: '1 1 130px' }} />
          <Select value={form.status} onChange={(e) => setForm({ ...form, status: e.target.value })} style={{ flex: '1 1 150px' }}>
            {ADVANCE_STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
          </Select>
        </div>
        <div style={{ marginBottom: 12 }}>
          <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Комментарий по авансу</label>
          <TextArea maxLength={200} rows={2} value={form.comment} onChange={(e) => setForm({ ...form, comment: e.target.value })} />
        </div>
        <button
          onClick={addAdvance}
          style={{ display: 'flex', alignItems: 'center', gap: 4, padding: '8px 14px', borderRadius: 7, border: 'none', background: COLORS.accent, color: '#fff', fontFamily: 'Inter, sans-serif', fontSize: 13, fontWeight: 500, cursor: 'pointer' }}
        >
          <Plus size={14} /> Добавить аванс
        </button>
      </BlueprintCard>

      {groups.map((group) => {
        const { key, label, items, sum } = group;
        const isCollapsed = collapsedGroups[key];
        return (
          <BlueprintCard key={key} style={{ marginBottom: 12, padding: 0, overflow: 'hidden' }}>
            <div
              onClick={() => setCollapsedGroups({ ...collapsedGroups, [key]: !isCollapsed })}
              style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer', padding: 16 }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <span style={{ display: 'inline-block', width: 10, height: 10, borderRadius: '50%', background: group.color }} />
                <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 15, fontWeight: 700, color: COLORS.ink }}>{label}</span>
                <span style={{ fontFamily: 'IBM Plex Mono, monospace', fontSize: 12, color: COLORS.inkSoft }}>{items.length} · {fmt(sum)}</span>
              </div>
              {isCollapsed ? <ChevronDown size={18} color={COLORS.inkSoft} /> : <ChevronUp size={18} color={COLORS.inkSoft} />}
            </div>

            {!isCollapsed && (
              items.length === 0 ? (
                <div style={{ padding: '0 16px 16px', fontFamily: 'Inter, sans-serif', fontSize: 13, color: COLORS.inkSoft, fontStyle: 'italic' }}>Пусто</div>
              ) : (
                <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                  <thead>
                    <tr>
                      <th style={{ ...thStyle, paddingLeft: 16 }}>Дата</th>
                      <th style={thStyle}>Сумма</th>
                      <th style={thStyle}>Статус</th>
                      <th style={{ ...thStyle, width: 40, paddingRight: 16 }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {items.map((a) => {
                      const isOpen = !!expandedIds[a.id];
                      const meta = ADVANCE_STATUS_META[a.status];
                      return (
                        <React.Fragment key={a.id}>
                          <tr onClick={() => toggleExpand(a.id)} style={{ cursor: 'pointer' }}>
                            <td style={{ ...tdStyle, paddingLeft: 16, fontFamily: 'IBM Plex Mono, monospace', fontSize: 12, color: COLORS.inkSoft }}>{formatDateRu(a.date)}</td>
                            <td style={{ ...tdStyle, fontFamily: 'IBM Plex Mono, monospace', fontSize: 13, color: COLORS.ink, fontWeight: 500 }}>{fmt(a.amount)}</td>
                            <td style={{ ...tdStyle }}>
                              <span style={{ fontFamily: 'Inter, sans-serif', fontSize: 12, fontWeight: 500, padding: '4px 10px', borderRadius: 999, background: meta.bg, color: meta.color }}>{a.status}</span>
                            </td>
                            <td style={{ ...tdStyle, paddingRight: 16, textAlign: 'right' }}>
                              {isOpen ? <ChevronUp size={16} color={COLORS.inkSoft} /> : <ChevronDown size={16} color={COLORS.inkSoft} />}
                            </td>
                          </tr>
                          {isOpen && (
                            <tr>
                              <td colSpan={4} style={{ padding: '0 16px 16px', borderBottom: `1px solid ${COLORS.bg}`, background: COLORS.bg }}>
                                {editingIds[a.id] ? (
                                  <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', paddingTop: 12 }}>
                                    <div style={{ flex: '1 1 150px' }}>
                                      <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Дата</label>
                                      <TextInput type="date" value={a.date} onChange={(ev) => updateField(a.id, 'date', ev.target.value)} />
                                    </div>
                                    <div style={{ flex: '1 1 130px' }}>
                                      <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Сумма, ₸</label>
                                      <TextInput type="number" value={a.amount} onChange={(ev) => updateField(a.id, 'amount', ev.target.value)} />
                                    </div>
                                    <div style={{ flex: '1 1 150px' }}>
                                      <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Статус</label>
                                      <Select value={a.status} onChange={(ev) => updateField(a.id, 'status', ev.target.value)}>
                                        {ADVANCE_STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
                                      </Select>
                                    </div>
                                    <div style={{ flex: '1 1 220px' }}>
                                      <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Комментарий по авансу</label>
                                      <TextArea maxLength={200} rows={2} value={a.comment || ''} onChange={(ev) => updateField(a.id, 'comment', ev.target.value)} />
                                    </div>
                                  </div>
                                ) : (
                                  <div style={{ paddingTop: 12 }}>
                                    <label style={{ fontSize: 11, color: COLORS.inkSoft }}>Комментарий по авансу</label>
                                    <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 13, color: COLORS.ink, padding: '6px 0' }}>{a.comment || '—'}</div>
                                  </div>
                                )}

                                <div style={{ marginTop: 12, display: 'flex', gap: 8 }}>
                                  <button
                                    onClick={() => toggleEdit(a.id)}
                                    style={{ display: 'flex', alignItems: 'center', gap: 4, padding: '6px 10px', borderRadius: 7, border: `1px solid ${editingIds[a.id] ? COLORS.accent : COLORS.line}`, background: editingIds[a.id] ? COLORS.accentSoft : COLORS.surface, color: editingIds[a.id] ? COLORS.accent : COLORS.ink, fontFamily: 'Inter, sans-serif', fontSize: 12, fontWeight: 500, cursor: 'pointer' }}
                                  >
                                    {editingIds[a.id] ? <><Check size={13} /> Готово</> : <><Pencil size={13} /> Редактировать</>}
                                  </button>
                                  <button
                                    onClick={() => removeAdvance(a.id)}
                                    style={{ display: 'flex', alignItems: 'center', gap: 4, padding: '6px 10px', borderRadius: 7, border: `1px solid ${COLORS.line}`, background: COLORS.surface, color: '#B23A2C', fontFamily: 'Inter, sans-serif', fontSize: 12, cursor: 'pointer' }}
                                  >
                                    <Trash2 size={13} /> Удалить аванс
                                  </button>
                                </div>
                              </td>
                            </tr>
                          )}
                        </React.Fragment>
                      );
                    })}
                  </tbody>
                </table>
              )
            )}
          </BlueprintCard>
        );
      })}
    </div>
  );
}

function Documents({ documents, setDocuments }) {
  const emptyForm = { name: '', date: '', url: '' };
  const [form, setForm] = useState(emptyForm);
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState({});

  const addDocument = () => {
    if (!form.name.trim() || !form.url.trim()) return;
    setDocuments((prev) => [...prev, { id: uid('d'), name: form.name.trim(), date: form.date, url: form.url.trim() }]);
    setForm(emptyForm);
  };

  const removeDocument = (id) => setDocuments((prev) => prev.filter((d) => d.id !== id));
  const startEdit = (d) => { setEditingId(d.id); setEditForm({ name: d.name, date: d.date, url: d.url }); };
  const saveEdit = (id) => {
    setDocuments((prev) => prev.map((d) => (d.id === id ? { ...d, name: editForm.name.trim() || d.name, date: editForm.date, url: editForm.url.trim() } : d)));
    setEditingId(null);
  };

  const sorted = useMemo(() => {
    return [...documents].sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  }, [documents]);

  return (
    <div>
      <BlueprintCard style={{ marginBottom: 20 }}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 15, fontWeight: 700, color: COLORS.ink, marginBottom: 12 }}>
          Новый документ
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          <TextInput placeholder="Название" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} style={{ flex: '2 1 200px' }} />
          <TextInput type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} style={{ flex: '1 1 150px' }} />
          <TextInput placeholder="Ссылка (https://…)" value={form.url} onChange={(e) => setForm({ ...form, url: e.target.value })} style={{ flex: '2 1 220px' }} />
          <button
            onClick={addDocument}
            style={{ display: 'flex', alignItems: 'center', gap: 4, padding: '8px 14px', borderRadius: 7, border: 'none', background: COLORS.accent, color: '#fff', fontFamily: 'Inter, sans-serif', fontSize: 13, fontWeight: 500, cursor: 'pointer' }}
          >
            <Plus size={14} /> Добавить документ
          </button>
        </div>
      </BlueprintCard>

      <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        {sorted.length === 0 && (
          <div style={{ fontFamily: 'Inter, sans-serif', fontSize: 13, color: COLORS.inkSoft, fontStyle: 'italic', textAlign: 'center', padding: 20 }}>
            Пока нет ни одного документа
          </div>
        )}
        {sorted.map((d) => {
          const isEditing = editingId === d.id;
          return (
            <BlueprintCard key={d.id}>
              {isEditing ? (
                <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
                  <TextInput placeholder="Название" value={editForm.name} onChange={(e) => setEditForm({ ...editForm, name: e.target.value })} style={{ flex: '2 1 180px' }} />
                  <TextInput type="date" value={editForm.date} onChange={(e) => setEditForm({ ...editForm, date: e.target.value })} style={{ flex: '1 1 150px' }} />
                  <TextInput placeholder="Ссылка" value={editForm.url} onChange={(e) => setEditForm({ ...editForm, url: e.target.value })} style={{ flex: '2 1 200px' }} />
                  <IconBtn onClick={() => saveEdit(d.id)} title="Сохранить"><Check size={14} /></IconBtn>
                  <IconBtn onClick={() => setEditingId(null)} title="Отмена"><X size={14} /></IconBtn>
                </div>
              ) : (
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 10, minWidth: 0 }}>
                    <Link2 size={16} color={COLORS.accent} style={{ flexShrink: 0 }} />
                    <div style={{ minWidth: 0 }}>
                      <a href={d.url} target="_blank" rel="noopener noreferrer" style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 14, fontWeight: 700, color: COLORS.ink, textDecoration: 'none' }}>
                        {d.name}
                      </a>
                      <div style={{ fontFamily: 'IBM Plex Mono, monospace', fontSize: 11, color: COLORS.inkSoft }}>{formatDateRu(d.date)}</div>
                    </div>
                  </div>
                  <div style={{ display: 'flex', gap: 6, flexShrink: 0 }}>
                    <IconBtn onClick={() => startEdit(d)} title="Редактировать"><Pencil size={14} /></IconBtn>
                    <IconBtn onClick={() => removeDocument(d.id)} title="Удалить" danger><Trash2 size={14} /></IconBtn>
                  </div>
                </div>
              )}
            </BlueprintCard>
          );
        })}
      </div>
    </div>
  );
}
