import { useState, useEffect, useCallback, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./lib/supabaseClient";
import {
  Gem, Mountain, LayoutDashboard, Search, Plus, Filter, Printer, Download,
  FileText, Trash2, Edit, Copy, Eye, QrCode, Users, History, Database,
  LogOut, Menu, X, Save, Upload, ChevronRight, ChevronLeft, AlertCircle,
  CheckCircle, BarChart3, Package, Calendar, MapPin, Lock, Shield, FileDown,
  FileUp, RefreshCw, Image as ImageIcon, ClipboardList, UserCog, Boxes,
  ArrowUpRight, TriangleAlert, ScanLine, FolderOpen
} from "lucide-react";

// ============================= CONSTANTES =============================
const GRUPOS_MINERALOGICOS = ["Silicatos", "Óxidos", "Sulfetos", "Carbonatos", "Sulfatos", "Haletos", "Fosfatos", "Elementos Nativos", "Boratos", "Outros"];
const CLASSES_QUIMICAS = ["Tectossilicatos", "Filossilicatos", "Inossilicatos", "Ciclossilicatos", "Nesossilicatos", "Sorossilicatos", "Óxidos Simples", "Óxidos Múltiplos", "Sulfetos Simples", "Carbonatos Anidros", "Outra"];
const SISTEMAS_CRISTALINOS = ["Cúbico", "Tetragonal", "Ortorrômbico", "Hexagonal", "Trigonal", "Monoclínico", "Triclínico", "Amorfo"];
const BRILHOS = ["Vítreo", "Metálico", "Adamantino", "Resinoso", "Gorduroso", "Sedoso", "Nacarado", "Fosco", "Terroso"];
const TRANSPARENCIAS = ["Transparente", "Translúcido", "Opaco"];
const CLIVAGENS = ["Perfeita", "Boa", "Regular", "Imperfeita", "Ausente"];
const FRATURAS = ["Concoidal", "Irregular", "Fibrosa", "Ganchosa", "Terrosa"];
const TENACIDADES = ["Frágil", "Maleável", "Séctil", "Dúctil", "Flexível", "Elástico"];
const ESTADOS_CONSERVACAO = ["Excelente", "Bom", "Regular", "Danificado"];
const TIPOS_ROCHA = ["Ígnea", "Sedimentar", "Metamórfica"];
const GRANULACOES = ["Afanítica (fina)", "Fanerítica (grossa)", "Vítrea", "Porfirítica", "Criptocristalina"];
const TEXTURAS_ROCHA = ["Granular", "Foliada", "Bandada", "Maciça", "Vesicular", "Clástica"];
const ESTADOS_BR = ["AC","AL","AP","AM","BA","CE","DF","ES","GO","MA","MT","MS","MG","PA","PB","PR","PE","PI","RJ","RN","RS","RO","RR","SC","SP","SE","TO"];

const MOHS_REF = [
  { d: 1, m: "Talco" }, { d: 2, m: "Gipsita" }, { d: 3, m: "Calcita" }, { d: 4, m: "Fluorita" },
  { d: 5, m: "Apatita" }, { d: 6, m: "Ortoclásio" }, { d: 7, m: "Quartzo" }, { d: 8, m: "Topázio" },
  { d: 9, m: "Coríndon" }, { d: 10, m: "Diamante" }
];

const emptyMineral = () => ({
  id: "", codigo: "", nome: "", nomeCientifico: "", grupo: "", classeQuimica: "",
  cor: "", traco: "", brilho: "", transparencia: "", clivagem: "", fratura: "",
  habito: "", sistemaCristalino: "", dureza: "", densidade: "", magnetismo: "Não",
  fluorescencia: "Não", tenacidade: "",
  composicaoQuimica: "", ambienteFormacao: "", usos: "", ocorrencias: "", estadoConservacao: "",
  localArmazenamento: "", armario: "", gaveta: "", prateleira: "", caixa: "",
  dataCadastro: "", responsavel: "", observacoes: "",
  fotoPrincipal: "", fotosAdicionais: [], arquivos: [],
  tipoAmostra: "mineral"
});

const emptyRocha = () => ({
  id: "", codigo: "", nome: "", tipoRocha: "",
  cor: "", granulacao: "", textura: "", estrutura: "", mineralogiaPrincipal: "", mineralogiaSecundaria: "",
  ambienteFormacao: "", idadeGeologica: "", origem: "",
  usoIndustrial: "", usoOrnamental: "", usoConstrucao: "",
  procedencia: "", municipio: "", estado: "", pais: "Brasil",
  localFisico: "", armario: "", gaveta: "", prateleira: "", caixa: "",
  responsavel: "", dataCadastro: "", observacoes: "",
  fotos: [],
  tipoAmostra: "rocha"
});

const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
const todayISO = () => new Date().toISOString().slice(0, 10);
const fmtDate = (iso) => { if (!iso) return "—"; const [y, m, d] = iso.split("-"); return `${d}/${m}/${y}`; };
const escXML = (s) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

function proximoCodigo(lista, prefixo) {
  let max = 0;
  lista.forEach((it) => {
    const m = (it.codigo || "").match(new RegExp(`^${prefixo}(\\d+)$`));
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return `${prefixo}${String(max + 1).padStart(3, "0")}`;
}

function downloadBlob(content, filename, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

function resizeImage(file, maxW = 640, quality = 0.7) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const scale = Math.min(1, maxW / img.width);
        const canvas = document.createElement("canvas");
        canvas.width = img.width * scale;
        canvas.height = img.height * scale;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.onerror = reject;
      img.src = e.target.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function mineralToXML(m) {
  return `  <mineral>
    <codigo>${escXML(m.codigo)}</codigo>
    <nome>${escXML(m.nome)}</nome>
    <nomeCientifico>${escXML(m.nomeCientifico)}</nomeCientifico>
    <grupo>${escXML(m.grupo)}</grupo>
    <classeQuimica>${escXML(m.classeQuimica)}</classeQuimica>
    <cor>${escXML(m.cor)}</cor>
    <traco>${escXML(m.traco)}</traco>
    <brilho>${escXML(m.brilho)}</brilho>
    <transparencia>${escXML(m.transparencia)}</transparencia>
    <clivagem>${escXML(m.clivagem)}</clivagem>
    <fratura>${escXML(m.fratura)}</fratura>
    <habitoCristalino>${escXML(m.habito)}</habitoCristalino>
    <sistemaCristalino>${escXML(m.sistemaCristalino)}</sistemaCristalino>
    <dureza>${escXML(m.dureza)}</dureza>
    <densidade>${escXML(m.densidade)}</densidade>
    <magnetismo>${escXML(m.magnetismo)}</magnetismo>
    <fluorescencia>${escXML(m.fluorescencia)}</fluorescencia>
    <tenacidade>${escXML(m.tenacidade)}</tenacidade>
    <composicaoQuimica>${escXML(m.composicaoQuimica)}</composicaoQuimica>
    <ambienteFormacao>${escXML(m.ambienteFormacao)}</ambienteFormacao>
    <principaisUsos>${escXML(m.usos)}</principaisUsos>
    <principaisOcorrencias>${escXML(m.ocorrencias)}</principaisOcorrencias>
    <estadoConservacao>${escXML(m.estadoConservacao)}</estadoConservacao>
    <localizacao>
        <local>${escXML(m.localArmazenamento)}</local>
        <armario>${escXML(m.armario)}</armario>
        <gaveta>${escXML(m.gaveta)}</gaveta>
        <prateleira>${escXML(m.prateleira)}</prateleira>
        <caixa>${escXML(m.caixa)}</caixa>
    </localizacao>
    <registro>
        <dataCadastro>${escXML(m.dataCadastro)}</dataCadastro>
        <responsavel>${escXML(m.responsavel)}</responsavel>
        <observacoes>${escXML(m.observacoes)}</observacoes>
    </registro>
  </mineral>`;
}

function rochaToXML(r) {
  return `  <rocha>
    <codigo>${escXML(r.codigo)}</codigo>
    <nome>${escXML(r.nome)}</nome>
    <tipo>${escXML(r.tipoRocha)}</tipo>
    <cor>${escXML(r.cor)}</cor>
    <granulacao>${escXML(r.granulacao)}</granulacao>
    <textura>${escXML(r.textura)}</textura>
    <estrutura>${escXML(r.estrutura)}</estrutura>
    <mineralogiaPrincipal>${escXML(r.mineralogiaPrincipal)}</mineralogiaPrincipal>
    <mineralogiaSecundaria>${escXML(r.mineralogiaSecundaria)}</mineralogiaSecundaria>
    <formacao>
        <ambienteFormacao>${escXML(r.ambienteFormacao)}</ambienteFormacao>
        <idadeGeologica>${escXML(r.idadeGeologica)}</idadeGeologica>
        <origem>${escXML(r.origem)}</origem>
    </formacao>
    <aplicacoes>
        <usoIndustrial>${escXML(r.usoIndustrial)}</usoIndustrial>
        <usoOrnamental>${escXML(r.usoOrnamental)}</usoOrnamental>
        <usoConstrucao>${escXML(r.usoConstrucao)}</usoConstrucao>
    </aplicacoes>
    <localizacaoOrigem>
        <procedencia>${escXML(r.procedencia)}</procedencia>
        <municipio>${escXML(r.municipio)}</municipio>
        <estado>${escXML(r.estado)}</estado>
        <pais>${escXML(r.pais)}</pais>
    </localizacaoOrigem>
    <controleFisico>
        <local>${escXML(r.localFisico)}</local>
        <armario>${escXML(r.armario)}</armario>
        <gaveta>${escXML(r.gaveta)}</gaveta>
        <prateleira>${escXML(r.prateleira)}</prateleira>
        <caixa>${escXML(r.caixa)}</caixa>
    </controleFisico>
    <registro>
        <dataCadastro>${escXML(r.dataCadastro)}</dataCadastro>
        <responsavel>${escXML(r.responsavel)}</responsavel>
        <observacoes>${escXML(r.observacoes)}</observacoes>
    </registro>
  </rocha>`;
}

// ============================= APP =============================
export default function AcervoGeologia() {
  const [booting, setBooting] = useState(true);
  const [storageError, setStorageError] = useState(false);
  const [auth, setAuth] = useState(null); // {username(email), nome, role, id}
  const [loginForm, setLoginForm] = useState({ u: "", p: "" });
  const [loginErr, setLoginErr] = useState("");
  const [authView, setAuthView] = useState("login"); // "login" | "register" | "forgot"
  const [regForm, setRegForm] = useState({ nome: "", username: "", password: "", confirm: "" });
  const [regErr, setRegErr] = useState("");
  const [regOk, setRegOk] = useState(false);
  const [forgotEmail, setForgotEmail] = useState("");
  const [forgotErr, setForgotErr] = useState("");
  const [forgotOk, setForgotOk] = useState(false);
  const [recoveryMode, setRecoveryMode] = useState(false);
  const [newPass, setNewPass] = useState({ p1: "", p2: "" });
  const [newPassErr, setNewPassErr] = useState("");
  const [newPassOk, setNewPassOk] = useState(false);

  const [minerais, setMinerais] = useState([]);
  const [rochas, setRochas] = useState([]);
  const [usuarios, setUsuarios] = useState([]); // perfis (tabela profiles)
  const [historico, setHistorico] = useState([]);

  const [page, setPage] = useState("dashboard");
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [mobileNavOpen, setMobileNavOpen] = useState(false);

  const [globalSearch, setGlobalSearch] = useState("");
  const [showGlobalResults, setShowGlobalResults] = useState(false);

  const [mineralForm, setMineralForm] = useState(null); // null = closed
  const [rochaForm, setRochaForm] = useState(null);
  const [viewItem, setViewItem] = useState(null); // ficha view/print
  const [toast, setToast] = useState(null);

  const printRef = useRef(null);

  // ---------- Carrega perfil (tabela profiles) a partir da sessão do Supabase Auth ----------
  const carregarPerfil = useCallback(async (user) => {
    const { data: perfil, error } = await supabase.from("profiles").select("*").eq("id", user.id).maybeSingle();
    if (error || !perfil) {
      setLoginErr("Não foi possível carregar seu perfil. Tente novamente.");
      await supabase.auth.signOut();
      return false;
    }
    if (perfil.status === "Pendente") {
      setLoginErr("Sua conta ainda está aguardando aprovação do administrador.");
      await supabase.auth.signOut();
      return false;
    }
    if (perfil.status === "Recusado") {
      setLoginErr("Seu cadastro foi recusado. Fale com o administrador.");
      await supabase.auth.signOut();
      return false;
    }
    setAuth({ username: perfil.email, nome: perfil.nome, role: perfil.role, id: perfil.id });
    return true;
  }, []);

  // ---------- Storage helpers (minerais / rochas / histórico continuam no kv_store) ----------
  const loadAll = useCallback(async () => {
    try {
      const [m, r, h] = await Promise.all([
        window.storage.get("minerais", true).catch(() => null),
        window.storage.get("rochas", true).catch(() => null),
        window.storage.get("historico", true).catch(() => null),
      ]);
      setMinerais(m ? JSON.parse(m.value) : []);
      setRochas(r ? JSON.parse(r.value) : []);
      setHistorico(h ? JSON.parse(h.value) : []);
    } catch (e) {
      setStorageError(true);
    }

    // Restaura sessão de autenticação já existente (Supabase Auth)
    const { data: { session } } = await supabase.auth.getSession();
    if (session?.user) await carregarPerfil(session.user);
    setBooting(false);
  }, [carregarPerfil]);

  useEffect(() => { loadAll(); }, [loadAll]);

  // Mantém a sessão sincronizada (ex: token expirado, logout em outra aba)
  useEffect(() => {
    const { data: listener } = supabase.auth.onAuthStateChange((event) => {
      if (event === "SIGNED_OUT") setAuth(null);
      if (event === "PASSWORD_RECOVERY") setRecoveryMode(true);
    });
    return () => listener.subscription.unsubscribe();
  }, []);

  // Carrega a lista de perfis (usuários) sempre que o administrador entra ou volta à aba
  const carregarPerfis = useCallback(async () => {
    const { data, error } = await supabase.from("profiles").select("*").order("created_at", { ascending: false });
    if (!error) setUsuarios(data || []);
  }, []);

  useEffect(() => { if (auth?.role === "Administrador") carregarPerfis(); }, [auth, carregarPerfis]);

  const persist = useCallback(async (key, value) => {
    try { await window.storage.set(key, JSON.stringify(value), true); }
    catch (e) { showToast("Falha ao salvar no banco de dados.", "error"); }
  }, []);

  const showToast = (msg, type = "success") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3200);
  };

  const logHistory = useCallback((acao, tipo, codigo, detalhes) => {
    setHistorico((prev) => {
      const entry = { id: uid(), data: new Date().toISOString(), usuario: auth?.nome || "—", acao, tipo, codigo, detalhes };
      const next = [entry, ...prev].slice(0, 500);
      persist("historico", next);
      return next;
    });
  }, [auth, persist]);

  // ---------- Auth (Supabase Auth) ----------
  const traduzErroAuth = (msg = "") => {
    if (/invalid login credentials/i.test(msg)) return "E-mail ou senha inválidos.";
    if (/email not confirmed/i.test(msg)) return "Confirme seu e-mail antes de entrar (verifique sua caixa de entrada).";
    if (/already registered|already exists/i.test(msg)) return "Já existe uma conta com esse e-mail.";
    if (/password should be/i.test(msg)) return "A senha não atende aos requisitos mínimos do Supabase.";
    return msg || "Ocorreu um erro. Tente novamente.";
  };

  const doLogin = async (e) => {
    if (e && e.preventDefault) e.preventDefault();
    setLoginErr("");
    if (!loginForm.u.trim() || !loginForm.p) { setLoginErr("Preencha e-mail e senha."); return; }
    const { data, error } = await supabase.auth.signInWithPassword({ email: loginForm.u.trim(), password: loginForm.p });
    if (error) { setLoginErr(traduzErroAuth(error.message)); return; }
    const ok = await carregarPerfil(data.user);
    if (ok) { setLoginErr(""); setPage("dashboard"); }
  };

  const doRegister = async () => {
    setRegErr("");
    if (!regForm.nome.trim() || !regForm.username.trim() || !regForm.password) { setRegErr("Preencha todos os campos."); return; }
    if (regForm.password.length < 6) { setRegErr("A senha deve ter pelo menos 6 caracteres."); return; }
    if (regForm.password !== regForm.confirm) { setRegErr("As senhas não coincidem."); return; }
    const email = regForm.username.trim();
    // O perfil (tabela "profiles") é criado automaticamente por um gatilho no banco
    // assim que o usuário é criado no Supabase Auth — ver supabase-schema.sql
    const { error } = await supabase.auth.signUp({ email, password: regForm.password, options: { data: { nome: regForm.nome.trim() } } });
    if (error) { setRegErr(traduzErroAuth(error.message)); return; }
    // Garante que a sessão criada automaticamente pelo signUp não fique logada antes da aprovação
    await supabase.auth.signOut();
    setRegOk(true);
    setRegForm({ nome: "", username: "", password: "", confirm: "" });
  };

  const logout = async () => { await supabase.auth.signOut(); setAuth(null); setLoginForm({ u: "", p: "" }); setPage("dashboard"); };

  const doForgotPassword = async () => {
    setForgotErr("");
    if (!forgotEmail.trim()) { setForgotErr("Informe seu e-mail."); return; }
    const { error } = await supabase.auth.resetPasswordForEmail(forgotEmail.trim(), { redirectTo: window.location.origin });
    if (error) { setForgotErr(traduzErroAuth(error.message)); return; }
    setForgotOk(true);
  };

  const doSetNewPassword = async () => {
    setNewPassErr("");
    if (!newPass.p1 || newPass.p1.length < 6) { setNewPassErr("A senha deve ter pelo menos 6 caracteres."); return; }
    if (newPass.p1 !== newPass.p2) { setNewPassErr("As senhas não coincidem."); return; }
    const { error } = await supabase.auth.updateUser({ password: newPass.p1 });
    if (error) { setNewPassErr(traduzErroAuth(error.message)); return; }
    setNewPassOk(true);
    setNewPass({ p1: "", p2: "" });
    await supabase.auth.signOut();
    setTimeout(() => { setRecoveryMode(false); setNewPassOk(false); setAuthView("login"); }, 2500);
  };
  const isAdmin = auth?.role === "Administrador";

  // ---------- CRUD Minerais ----------
  const saveMineral = async (data, isDuplicate = false) => {
    if (!isAdmin) return showToast("Apenas administradores podem cadastrar ou editar minerais.", "error");
    let next;
    const isNew = !data.id || isDuplicate;
    if (isNew) {
      const novo = { ...data, id: uid(), codigo: isDuplicate ? proximoCodigo(minerais, "MIN") : (data.codigo || proximoCodigo(minerais, "MIN")), dataCadastro: data.dataCadastro || todayISO(), responsavel: data.responsavel || auth.nome };
      next = [novo, ...minerais];
      logHistory(isDuplicate ? "Duplicado" : "Cadastrado", "Mineral", novo.codigo, novo.nome);
    } else {
      next = minerais.map((m) => (m.id === data.id ? data : m));
      logHistory("Editado", "Mineral", data.codigo, data.nome);
    }
    setMinerais(next);
    await persist("minerais", next);
    setMineralForm(null);
    showToast(isNew ? "Mineral cadastrado com sucesso." : "Mineral atualizado com sucesso.");
  };

  const deleteMineral = async (item) => {
    if (!isAdmin) return showToast("Apenas administradores podem excluir registros.", "error");
    if (!confirm(`Excluir o mineral "${item.nome}" (${item.codigo})? Esta ação não pode ser desfeita.`)) return;
    const next = minerais.filter((m) => m.id !== item.id);
    setMinerais(next);
    await persist("minerais", next);
    logHistory("Excluído", "Mineral", item.codigo, item.nome);
    showToast("Mineral excluído.");
  };

  // ---------- CRUD Rochas ----------
  const saveRocha = async (data, isDuplicate = false) => {
    if (!isAdmin) return showToast("Apenas administradores podem cadastrar ou editar rochas.", "error");
    let next;
    const isNew = !data.id || isDuplicate;
    if (isNew) {
      const novo = { ...data, id: uid(), codigo: isDuplicate ? proximoCodigo(rochas, "ROC") : (data.codigo || proximoCodigo(rochas, "ROC")), dataCadastro: data.dataCadastro || todayISO(), responsavel: data.responsavel || auth.nome };
      next = [novo, ...rochas];
      logHistory(isDuplicate ? "Duplicado" : "Cadastrado", "Rocha", novo.codigo, novo.nome);
    } else {
      next = rochas.map((r) => (r.id === data.id ? data : r));
      logHistory("Editado", "Rocha", data.codigo, data.nome);
    }
    setRochas(next);
    await persist("rochas", next);
    setRochaForm(null);
    showToast(isNew ? "Rocha cadastrada com sucesso." : "Rocha atualizada com sucesso.");
  };

  const deleteRocha = async (item) => {
    if (!isAdmin) return showToast("Apenas administradores podem excluir registros.", "error");
    if (!confirm(`Excluir a rocha "${item.nome}" (${item.codigo})? Esta ação não pode ser desfeita.`)) return;
    const next = rochas.filter((r) => r.id !== item.id);
    setRochas(next);
    await persist("rochas", next);
    logHistory("Excluído", "Rocha", item.codigo, item.nome);
    showToast("Rocha excluída.");
  };

  // ---------- Usuários (tabela profiles no Supabase) ----------
  const saveUsuario = async (data) => {
    if (!isAdmin) return showToast("Apenas administradores podem gerenciar usuários.", "error");
    const { error } = await supabase.from("profiles").update({ nome: data.nome, role: data.role, status: data.status }).eq("id", data.id);
    if (error) return showToast("Erro ao salvar usuário: " + error.message, "error");
    await carregarPerfis();
    logHistory(`Perfil atualizado (${data.status}/${data.role})`, "Usuário", "-", data.nome);
    showToast("Usuário salvo com sucesso.");
  };

  const revogarUsuario = async (u) => {
    if (!isAdmin) return showToast("Apenas administradores podem gerenciar usuários.", "error");
    if (u.id === auth.id) return showToast("Você não pode revogar seu próprio acesso.", "error");
    if (!confirm(`Revogar o acesso de "${u.nome}"? A conta deixa de conseguir entrar no sistema.`)) return;
    const { error } = await supabase.from("profiles").update({ status: "Recusado" }).eq("id", u.id);
    if (error) return showToast("Erro ao revogar acesso: " + error.message, "error");
    await carregarPerfis();
    showToast("Acesso revogado.");
  };

  // ---------- Export helpers ----------
  const exportXMLItem = (item) => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>\n<acervoGeologico>\n${item.tipoAmostra === "mineral" ? mineralToXML(item) : rochaToXML(item)}\n</acervoGeologico>`;
    downloadBlob(xml, `${item.codigo}.xml`, "application/xml");
    logHistory("Exportado XML", item.tipoAmostra === "mineral" ? "Mineral" : "Rocha", item.codigo, item.nome);
  };

  const exportXMLAll = () => {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>\n<acervoGeologico>\n${minerais.map(mineralToXML).join("\n")}\n${rochas.map(rochaToXML).join("\n")}\n</acervoGeologico>`;
    downloadBlob(xml, `acervo_completo_${todayISO()}.xml`, "application/xml");
    showToast("Exportação XML completa gerada.");
  };

  const exportExcel = (rows, sheetName, filename) => {
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, filename);
  };

  const printFicha = (item) => {
    setViewItem(item);
    setTimeout(() => window.print(), 250);
  };

  const backupJSON = () => {
    const data = { minerais, rochas, usuarios, historico, exportadoEm: new Date().toISOString(), versao: "1.0" };
    downloadBlob(JSON.stringify(data, null, 2), `backup_acervo_${todayISO()}.json`, "application/json");
    showToast("Backup gerado com sucesso.");
  };

  const restoreJSON = async (file) => {
    if (!isAdmin) return showToast("Apenas administradores podem restaurar backups.", "error");
    try {
      const text = await file.text();
      const data = JSON.parse(text);
      if (!confirm("Restaurar este backup substituirá TODOS os dados atuais. Deseja continuar?")) return;
      if (data.minerais) { setMinerais(data.minerais); await persist("minerais", data.minerais); }
      if (data.rochas) { setRochas(data.rochas); await persist("rochas", data.rochas); }
      if (data.usuarios) { setUsuarios(data.usuarios); await persist("usuarios", data.usuarios); }
      if (data.historico) { setHistorico(data.historico); await persist("historico", data.historico); }
      logHistory("Restauração de backup", "Sistema", "-", `Backup de ${data.exportadoEm || "data desconhecida"}`);
      showToast("Backup restaurado com sucesso.");
    } catch (e) { showToast("Arquivo de backup inválido.", "error"); }
  };

  // ---------- Derived ----------
  const allItems = useMemo(() => [...minerais, ...rochas], [minerais, rochas]);
  const stats = useMemo(() => {
    const now = new Date(); const ym = now.toISOString().slice(0, 7);
    const doMes = allItems.filter((i) => (i.dataCadastro || "").startsWith(ym)).length;
    return { totalMinerais: minerais.length, totalRochas: rochas.length, doMes, total: allItems.length };
  }, [minerais, rochas, allItems]);

  const ultimosCadastros = useMemo(() =>
    [...allItems].sort((a, b) => (b.dataCadastro || "").localeCompare(a.dataCadastro || "")).slice(0, 6),
    [allItems]);

  const globalResults = useMemo(() => {
    const q = globalSearch.trim().toLowerCase();
    if (!q) return [];
    return allItems.filter((i) => {
      const local = i.tipoAmostra === "mineral" ? i.localArmazenamento : i.localFisico;
      const municipio = i.municipio || "";
      const classe = i.tipoAmostra === "mineral" ? i.classeQuimica : i.tipoRocha;
      return [i.nome, i.codigo, i.cor, classe, i.tipoAmostra, local, municipio]
        .some((v) => (v || "").toLowerCase().includes(q));
    }).slice(0, 12);
  }, [globalSearch, allItems]);

  // ============================= RENDER =============================
  if (booting) {
    return (
      <div style={S.bootWrap}>
        <style>{GLOBAL_CSS}</style>
        <div style={S.bootCard}>
          <div className="crystalSpin" style={S.crystal} />
          <p style={{ fontFamily: "Inter", color: "#0B3D2E", marginTop: 16, fontWeight: 600 }}>Carregando acervo geológico…</p>
        </div>
      </div>
    );
  }

  if (recoveryMode) {
    return (
      <div style={S.loginWrap}>
        <style>{GLOBAL_CSS}</style>
        <div style={S.loginBg} />
        <div style={S.loginCard}>
          <div style={S.loginBrand}>
            <div style={S.hexIcon}><Lock size={24} color="#F7F9F6" /></div>
            <div>
              <div style={{ fontFamily: "Fraunces", fontWeight: 600, fontSize: 22, color: "#0B3D2E", lineHeight: 1.1 }}>Nova senha</div>
              <div style={{ fontFamily: "Inter", fontSize: 12.5, color: "#5b7768", letterSpacing: 0.3 }}>GeoAcervo — redefinição de senha</div>
            </div>
          </div>
          {newPassOk ? (
            <div style={S.regOkBox}>
              <CheckCircle size={26} color="#2E7D5B" />
              <p style={{ fontWeight: 700, color: "#0B3D2E", margin: "10px 0 4px" }}>Senha redefinida!</p>
              <p style={{ fontSize: 12.5, color: "#5b7768", lineHeight: 1.6 }}>Você já pode entrar com sua nova senha.</p>
            </div>
          ) : (
            <div style={{ marginTop: 20 }}>
              <label style={S.label}>Nova senha</label>
              <input style={S.input} type="password" value={newPass.p1} onChange={(e) => setNewPass({ ...newPass, p1: e.target.value })} placeholder="Mínimo 6 caracteres" autoFocus />
              <label style={S.label}>Confirmar nova senha</label>
              <input style={S.input} type="password" value={newPass.p2} onChange={(e) => setNewPass({ ...newPass, p2: e.target.value })} placeholder="Repita a senha" onKeyDown={(e) => e.key === "Enter" && doSetNewPassword()} />
              {newPassErr && <div style={S.loginErr}><AlertCircle size={15} /> {newPassErr}</div>}
              <button type="button" onClick={doSetNewPassword} style={S.btnPrimary}>Salvar nova senha <ChevronRight size={16} /></button>
            </div>
          )}
        </div>
      </div>
    );
  }

  if (!auth) {
    return (
      <div style={S.loginWrap}>
        <style>{GLOBAL_CSS}</style>
        <div style={S.loginBg} />
        <div style={S.loginCard}>
          <div style={S.loginBrand}>
            <div style={S.hexIcon}><Gem size={26} color="#F7F9F6" /></div>
            <div>
              <div style={{ fontFamily: "Fraunces", fontWeight: 600, fontSize: 22, color: "#0B3D2E", lineHeight: 1.1 }}>GeoAcervo</div>
              <div style={{ fontFamily: "Inter", fontSize: 12.5, color: "#5b7768", letterSpacing: 0.3 }}>Gestão de Minerais & Rochas</div>
            </div>
          </div>
          {authView !== "forgot" && (
            <div style={S.authTabs}>
              <button type="button" style={{ ...S.authTab, ...(authView === "login" ? S.authTabActive : {}) }} onClick={() => { setAuthView("login"); setRegErr(""); setRegOk(false); }}>Entrar</button>
              <button type="button" style={{ ...S.authTab, ...(authView === "register" ? S.authTabActive : {}) }} onClick={() => { setAuthView("register"); setLoginErr(""); }}>Criar conta</button>
            </div>
          )}

          {authView === "login" ? (
            <>
              <div style={{ marginTop: 20 }}>
                <label style={S.label}>E-mail</label>
                <input style={S.input} type="email" value={loginForm.u} onChange={(e) => setLoginForm({ ...loginForm, u: e.target.value })} placeholder="seu@email.com" autoFocus onKeyDown={(e) => e.key === "Enter" && doLogin()} />
                <label style={S.label}>Senha</label>
                <input style={S.input} type="password" value={loginForm.p} onChange={(e) => setLoginForm({ ...loginForm, p: e.target.value })} placeholder="••••••••" onKeyDown={(e) => e.key === "Enter" && doLogin()} />
                {loginErr && <div style={S.loginErr}><AlertCircle size={15} /> {loginErr}</div>}
                <button type="button" onClick={doLogin} style={S.btnPrimary}>Entrar no sistema <ChevronRight size={16} /></button>
                <button type="button" style={S.forgotLink} onClick={() => { setAuthView("forgot"); setForgotErr(""); setForgotOk(false); setForgotEmail(loginForm.u); }}>Esqueci minha senha</button>
              </div>
            </>
          ) : authView === "forgot" ? (
            forgotOk ? (
              <div style={S.regOkBox}>
                <CheckCircle size={26} color="#2E7D5B" />
                <p style={{ fontWeight: 700, color: "#0B3D2E", margin: "10px 0 4px" }}>E-mail enviado!</p>
                <p style={{ fontSize: 12.5, color: "#5b7768", lineHeight: 1.6 }}>Se esse e-mail estiver cadastrado, você vai receber um link para redefinir sua senha. Confira também a caixa de spam.</p>
                <button type="button" style={S.btnGhost} onClick={() => { setAuthView("login"); setForgotOk(false); }}>Voltar para o login</button>
              </div>
            ) : (
              <div style={{ marginTop: 20 }}>
                <p style={{ fontSize: 12.5, color: "#5b7768", lineHeight: 1.6, marginBottom: 14 }}>Informe o e-mail da sua conta. Enviaremos um link para você definir uma nova senha.</p>
                <label style={S.label}>E-mail</label>
                <input style={S.input} type="email" value={forgotEmail} onChange={(e) => setForgotEmail(e.target.value)} placeholder="seu@email.com" autoFocus onKeyDown={(e) => e.key === "Enter" && doForgotPassword()} />
                {forgotErr && <div style={S.loginErr}><AlertCircle size={15} /> {forgotErr}</div>}
                <button type="button" onClick={doForgotPassword} style={S.btnPrimary}>Enviar link de recuperação <ChevronRight size={16} /></button>
                <button type="button" style={S.forgotLink} onClick={() => setAuthView("login")}>Voltar para o login</button>
              </div>
            )
          ) : regOk ? (
            <div style={S.regOkBox}>
              <CheckCircle size={26} color="#2E7D5B" />
              <p style={{ fontWeight: 700, color: "#0B3D2E", margin: "10px 0 4px" }}>Cadastro enviado!</p>
              <p style={{ fontSize: 12.5, color: "#5b7768", lineHeight: 1.6 }}>Sua conta está aguardando aprovação do administrador. Você poderá entrar assim que ela for aprovada (e, se a confirmação de e-mail estiver ativa no projeto, após confirmar seu e-mail).</p>
              <button type="button" style={S.btnGhost} onClick={() => { setAuthView("login"); setRegOk(false); }}>Voltar para o login</button>
            </div>
          ) : (
            <div style={{ marginTop: 20 }}>
              <label style={S.label}>Nome completo</label>
              <input style={S.input} value={regForm.nome} onChange={(e) => setRegForm({ ...regForm, nome: e.target.value })} placeholder="Seu nome" />
              <label style={S.label}>E-mail</label>
              <input style={S.input} type="email" value={regForm.username} onChange={(e) => setRegForm({ ...regForm, username: e.target.value })} placeholder="seu@email.com" />
              <label style={S.label}>Senha</label>
              <input style={S.input} type="password" value={regForm.password} onChange={(e) => setRegForm({ ...regForm, password: e.target.value })} placeholder="Mínimo 6 caracteres" />
              <label style={S.label}>Confirmar senha</label>
              <input style={S.input} type="password" value={regForm.confirm} onChange={(e) => setRegForm({ ...regForm, confirm: e.target.value })} placeholder="Repita a senha" onKeyDown={(e) => e.key === "Enter" && doRegister()} />
              {regErr && <div style={S.loginErr}><AlertCircle size={15} /> {regErr}</div>}
              <button type="button" onClick={doRegister} style={S.btnPrimary}>Solicitar cadastro <ChevronRight size={16} /></button>
              <p style={{ fontSize: 11.5, color: "#a9bdb2", marginTop: 12, lineHeight: 1.6 }}>Sua conta ficará com status "Pendente" até um administrador aprovar o acesso.</p>
            </div>
          )}
        </div>
      </div>
    );
  }

  return (
    <div style={S.app}>
      <style>{GLOBAL_CSS}</style>
      {storageError && (
        <div style={S.storageWarn}><TriangleAlert size={15} /> Não foi possível conectar ao banco de dados. Trabalhando em modo temporário.</div>
      )}

      {/* Sidebar */}
      <aside style={{ ...S.sidebar, width: sidebarOpen ? 240 : 76 }} className={mobileNavOpen ? "mobileNavOpen" : ""}>
        <div style={S.sidebarTop}>
          <div style={S.hexIconSm}><Gem size={18} color="#0B3D2E" /></div>
          {sidebarOpen && <span style={{ fontFamily: "Fraunces", fontWeight: 600, fontSize: 17, color: "#F7F9F6" }}>GeoAcervo</span>}
        </div>
        <nav style={S.nav}>
          <NavItem icon={LayoutDashboard} label="Dashboard" active={page === "dashboard"} open={sidebarOpen} onClick={() => { setPage("dashboard"); setMobileNavOpen(false); }} />
          <NavItem icon={Gem} label="Minerais" active={page === "minerais"} open={sidebarOpen} onClick={() => { setPage("minerais"); setMobileNavOpen(false); }} />
          <NavItem icon={Mountain} label="Rochas" active={page === "rochas"} open={sidebarOpen} onClick={() => { setPage("rochas"); setMobileNavOpen(false); }} />
          {isAdmin && <NavItem icon={ClipboardList} label="Controle de Cadastros" active={page === "controle"} open={sidebarOpen} onClick={() => { setPage("controle"); setMobileNavOpen(false); }} />}
          {isAdmin && <NavItem icon={BarChart3} label="Relatórios" active={page === "relatorios"} open={sidebarOpen} onClick={() => { setPage("relatorios"); setMobileNavOpen(false); }} />}
          <NavItem icon={History} label="Histórico" active={page === "historico"} open={sidebarOpen} onClick={() => { setPage("historico"); setMobileNavOpen(false); }} />
          {isAdmin && <NavItem icon={UserCog} label="Usuários" active={page === "usuarios"} open={sidebarOpen} onClick={() => { setPage("usuarios"); setMobileNavOpen(false); }} badge={usuarios.filter((u) => u.status === "Pendente").length} />}
          <NavItem icon={Database} label="Backup" active={page === "backup"} open={sidebarOpen} onClick={() => { setPage("backup"); setMobileNavOpen(false); }} />
        </nav>
        <div style={S.sidebarBottom}>
          <button style={S.sidebarCollapse} onClick={() => setSidebarOpen((v) => !v)}>{sidebarOpen ? <ChevronLeft size={16} /> : <ChevronRight size={16} />}</button>
          <button style={S.logoutBtn} onClick={logout}><LogOut size={16} /> {sidebarOpen && "Sair"}</button>
        </div>
      </aside>
      {mobileNavOpen && <div style={S.overlay} onClick={() => setMobileNavOpen(false)} />}

      {/* Main */}
      <div style={S.main}>
        <header style={S.topbar}>
          <button className="onlyMobile" style={S.menuBtn} onClick={() => setMobileNavOpen(true)}><Menu size={20} /></button>
          <div style={S.searchBox}>
            <Search size={16} color="#5b7768" />
            <input
              style={S.searchInput}
              placeholder="Pesquisa inteligente: nome, código, cor, classe, tipo, local, município…"
              value={globalSearch}
              onChange={(e) => { setGlobalSearch(e.target.value); setShowGlobalResults(true); }}
              onFocus={() => setShowGlobalResults(true)}
              onBlur={() => setTimeout(() => setShowGlobalResults(false), 180)}
            />
            {showGlobalResults && globalSearch && (
              <div style={S.searchResults}>
                {globalResults.length === 0 && <div style={S.searchEmpty}>Nenhuma amostra encontrada para "{globalSearch}".</div>}
                {globalResults.map((r) => (
                  <div key={r.id} style={S.searchResultRow} onMouseDown={() => { setViewItem(r); setGlobalSearch(""); }}>
                    {r.tipoAmostra === "mineral" ? <Gem size={14} color="#2E7D5B" /> : <Mountain size={14} color="#C9A227" />}
                    <span style={S.mono}>{r.codigo}</span>
                    <span style={{ fontWeight: 600 }}>{r.nome}</span>
                    <span style={{ color: "#7c9186", fontSize: 12.5 }}>{r.cor}</span>
                  </div>
                ))}
              </div>
            )}
          </div>
          <div style={S.userChip}>
            <div style={S.userAvatar}>{auth.nome.slice(0, 1).toUpperCase()}</div>
            <div style={{ lineHeight: 1.2 }}>
              <div style={{ fontWeight: 600, fontSize: 13.5 }}>{auth.nome}</div>
              <div style={{ fontSize: 11.5, color: "#5b7768" }}>{auth.role}</div>
            </div>
          </div>
        </header>

        <main style={S.content}>
          {page === "dashboard" && (
            <Dashboard stats={stats} ultimos={ultimosCadastros} onNovoMineral={() => setMineralForm(emptyMineral())} onNovaRocha={() => setRochaForm(emptyRocha())} onControle={() => setPage("controle")} onRelatorios={() => setPage("relatorios")} onView={setViewItem} minerais={minerais} rochas={rochas} isAdmin={isAdmin} />
          )}
          {page === "minerais" && (
            <ListaPage
              title="Minerais" icon={Gem} color="#2E7D5B" items={minerais}
              onNew={() => setMineralForm(emptyMineral())}
              onView={setViewItem}
              onEdit={(it) => setMineralForm(it)}
              onDuplicate={(it) => saveMineral({ ...it, id: "" }, true)}
              onDelete={deleteMineral}
              onPrint={printFicha}
              onExportXML={exportXMLItem}
              badgeField="grupo"
              isAdmin={isAdmin}
            />
          )}
          {page === "rochas" && (
            <ListaPage
              title="Rochas" icon={Mountain} color="#C9A227" items={rochas}
              onNew={() => setRochaForm(emptyRocha())}
              onView={setViewItem}
              onEdit={(it) => setRochaForm(it)}
              onDuplicate={(it) => saveRocha({ ...it, id: "" }, true)}
              onDelete={deleteRocha}
              onPrint={printFicha}
              onExportXML={exportXMLItem}
              badgeField="tipoRocha"
              isAdmin={isAdmin}
            />
          )}
          {page === "controle" && isAdmin && (
            <ControlePage
              allItems={allItems}
              onView={setViewItem}
              onEdit={(it) => it.tipoAmostra === "mineral" ? setMineralForm(it) : setRochaForm(it)}
              onDuplicate={(it) => it.tipoAmostra === "mineral" ? saveMineral({ ...it, id: "" }, true) : saveRocha({ ...it, id: "" }, true)}
              onDelete={(it) => it.tipoAmostra === "mineral" ? deleteMineral(it) : deleteRocha(it)}
              onPrint={printFicha}
              onExportXML={exportXMLItem}
              responsaveis={[...new Set(allItems.map((i) => i.responsavel).filter(Boolean))]}
            />
          )}
          {page === "relatorios" && isAdmin && (
            <RelatoriosPage minerais={minerais} rochas={rochas} exportExcel={exportExcel} exportXMLAll={exportXMLAll} />
          )}
          {(page === "controle" || page === "relatorios") && !isAdmin && <RestritoPage />}
          {page === "historico" && <HistoricoPage historico={historico} />}
          {page === "usuarios" && isAdmin && <UsuariosPage usuarios={usuarios} onSave={saveUsuario} onRevoke={revogarUsuario} authId={auth.id} />}
          {page === "backup" && <BackupPage onBackup={backupJSON} onRestore={restoreJSON} isAdmin={isAdmin} stats={stats} />}
        </main>
      </div>

      {mineralForm && <MineralFormModal data={mineralForm} onClose={() => setMineralForm(null)} onSave={saveMineral} auth={auth} />}
      {rochaForm && <RochaFormModal data={rochaForm} onClose={() => setRochaForm(null)} onSave={saveRocha} auth={auth} />}
      {viewItem && <FichaModal item={viewItem} onClose={() => setViewItem(null)} onEdit={(it) => { setViewItem(null); it.tipoAmostra === "mineral" ? setMineralForm(it) : setRochaForm(it); }} onExportXML={exportXMLItem} />}

      {toast && (
        <div style={{ ...S.toast, background: toast.type === "error" ? "#8a2f2f" : "#0B3D2E" }}>
          {toast.type === "error" ? <AlertCircle size={16} /> : <CheckCircle size={16} />} {toast.msg}
        </div>
      )}
    </div>
  );
}

// ============================= COMPONENTES =============================

function NavItem({ icon: Icon, label, active, onClick, open, badge }) {
  return (
    <button onClick={onClick} style={{ ...S.navItem, ...(active ? S.navItemActive : {}), justifyContent: open ? "flex-start" : "center" }} title={label}>
      <Icon size={18} />
      {open && <span style={{ flex: 1 }}>{label}</span>}
      {badge > 0 && <span style={S.navBadge}>{badge}</span>}
    </button>
  );
}

function StatCard({ icon: Icon, label, value, accent, sub }) {
  return (
    <div style={S.statCard}>
      <div style={{ ...S.statIcon, background: accent }}><Icon size={20} color="#fff" /></div>
      <div>
        <div style={S.statValue}>{value}</div>
        <div style={S.statLabel}>{label}</div>
        {sub && <div style={S.statSub}>{sub}</div>}
      </div>
    </div>
  );
}

function Dashboard({ stats, ultimos, onNovoMineral, onNovaRocha, onControle, onRelatorios, onView, minerais, rochas, isAdmin }) {
  return (
    <div>
      <PageHeader title="Painel Geral" subtitle="Visão consolidada do acervo de minerais e rochas do laboratório" />
      <div style={S.statsGrid}>
        <StatCard icon={Gem} label="Minerais cadastrados" value={stats.totalMinerais} accent="#2E7D5B" />
        <StatCard icon={Mountain} label="Rochas cadastradas" value={stats.totalRochas} accent="#C9A227" />
        <StatCard icon={Calendar} label="Cadastros neste mês" value={stats.doMes} accent="#0B3D2E" />
        <StatCard icon={Boxes} label="Total no acervo" value={stats.total} accent="#5b7768" sub="minerais + rochas" />
      </div>

      {isAdmin && (
        <div style={S.quickGrid}>
          <QuickAction icon={Plus} label="Novo Mineral" onClick={onNovoMineral} color="#2E7D5B" />
          <QuickAction icon={Plus} label="Nova Rocha" onClick={onNovaRocha} color="#C9A227" />
          <QuickAction icon={ClipboardList} label="Controle de Cadastros" onClick={onControle} color="#0B3D2E" />
          <QuickAction icon={BarChart3} label="Relatórios" onClick={onRelatorios} color="#5b7768" />
        </div>
      )}

      <div style={S.card}>
        <div style={S.cardHeadRow}>
          <h3 style={S.cardTitle}>Últimos cadastros</h3>
        </div>
        {ultimos.length === 0 ? (
          <EmptyState text="Nenhuma amostra cadastrada ainda. Use os atalhos acima para começar." />
        ) : (
          <div style={S.tableWrap}>
            <table style={S.table}>
              <thead><tr><th>Código</th><th>Nome</th><th>Tipo</th><th>Responsável</th><th>Data</th><th></th></tr></thead>
              <tbody>
                {ultimos.map((it) => (
                  <tr key={it.id}>
                    <td style={S.mono}>{it.codigo}</td>
                    <td style={{ fontWeight: 600 }}>{it.nome}</td>
                    <td><TypeBadge item={it} /></td>
                    <td>{it.responsavel}</td>
                    <td>{fmtDate(it.dataCadastro)}</td>
                    <td><button style={S.iconBtn} onClick={() => onView(it)}><Eye size={15} /></button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

function QuickAction({ icon: Icon, label, onClick, color }) {
  return (
    <button style={S.quickCard} onClick={onClick}>
      <div style={{ ...S.quickIcon, background: color }}><Icon size={18} color="#fff" /></div>
      <span>{label}</span>
      <ArrowUpRight size={15} color="#7c9186" style={{ marginLeft: "auto" }} />
    </button>
  );
}

function TypeBadge({ item }) {
  const isMin = item.tipoAmostra === "mineral";
  return (
    <span style={{ ...S.badge, background: isMin ? "#E4F2E9" : "#FBF2DA", color: isMin ? "#0B3D2E" : "#7a5f0a" }}>
      {isMin ? <Gem size={11} /> : <Mountain size={11} />} {isMin ? (item.grupo || "Mineral") : (item.tipoRocha || "Rocha")}
    </span>
  );
}

function PageHeader({ title, subtitle, right }) {
  return (
    <div style={S.pageHeader}>
      <div>
        <h1 style={S.pageTitle}>{title}</h1>
        {subtitle && <p style={S.pageSubtitle}>{subtitle}</p>}
      </div>
      {right}
    </div>
  );
}

function EmptyState({ text, icon: Icon = FolderOpen }) {
  return (
    <div style={S.emptyState}>
      <Icon size={30} color="#a9bdb2" />
      <p>{text}</p>
    </div>
  );
}

function RestritoPage() {
  return (
    <div style={S.card}>
      <div style={S.emptyState}>
        <Lock size={30} color="#a9bdb2" />
        <p>Esta área é exclusiva para o perfil Administrador.</p>
      </div>
    </div>
  );
}

function ListaPage({ title, icon: Icon, color, items, onNew, onView, onEdit, onDuplicate, onDelete, onPrint, onExportXML, badgeField, isAdmin }) {
  const [q, setQ] = useState("");
  const filtered = items.filter((i) => !q || [i.nome, i.codigo, i[badgeField]].some((v) => (v || "").toLowerCase().includes(q.toLowerCase())));
  return (
    <div>
      <PageHeader title={title} subtitle={`${items.length} registro(s) no acervo`} right={
        isAdmin && <button style={S.btnPrimarySm} onClick={onNew}><Plus size={16} /> Novo{title === "Rochas" ? "a" : ""} {title.slice(0, -1)}</button>
      } />
      <div style={S.filterRow}>
        <div style={S.searchBoxSm}><Search size={15} color="#5b7768" /><input style={S.searchInputSm} placeholder={`Buscar por nome, código ou ${badgeField === "grupo" ? "grupo" : "tipo"}…`} value={q} onChange={(e) => setQ(e.target.value)} /></div>
      </div>
      {filtered.length === 0 ? (
        <div style={S.card}><EmptyState text={items.length === 0 ? `Nenhum registro ainda. Clique em "Novo" para cadastrar a primeira amostra.` : "Nenhum resultado para essa busca."} icon={Icon} /></div>
      ) : (
        <div style={S.cardGrid}>
          {filtered.map((it) => (
            <SampleCard key={it.id} item={it} color={color} onView={onView} onEdit={onEdit} onDuplicate={onDuplicate} onDelete={onDelete} onPrint={onPrint} onExportXML={onExportXML} isAdmin={isAdmin} />
          ))}
        </div>
      )}
    </div>
  );
}

function SampleCard({ item, color, onView, onEdit, onDuplicate, onDelete, onPrint, onExportXML, isAdmin }) {
  const isMin = item.tipoAmostra === "mineral";
  const foto = isMin ? item.fotoPrincipal : item.fotos?.[0];
  return (
    <div style={S.sampleCard}>
      <div style={S.sampleImg}>
        {foto ? <img src={foto} alt={item.nome} style={S.sampleImgTag} /> : <div style={S.sampleImgPh}><ImageIcon size={22} color="#a9bdb2" /></div>}
        <span style={{ ...S.codeChip, borderColor: color }}>{item.codigo}</span>
      </div>
      <div style={{ padding: "12px 14px 14px" }}>
        <div style={{ fontFamily: "Fraunces", fontWeight: 600, fontSize: 16.5, color: "#0B3D2E" }}>{item.nome || "(sem nome)"}</div>
        <div style={{ fontSize: 12.5, color: "#5b7768", marginTop: 2 }}>{isMin ? (item.grupo || "—") : (item.tipoRocha || "—")} · {item.cor || "sem cor definida"}</div>
        <div style={{ display: "flex", gap: 6, marginTop: 12, flexWrap: "wrap" }}>
          <button style={S.iconBtn} title="Visualizar" onClick={() => onView(item)}><Eye size={14} /></button>
          <button style={S.iconBtn} title="Imprimir ficha" onClick={() => onPrint(item)}><Printer size={14} /></button>
          <button style={S.iconBtn} title="Exportar XML" onClick={() => onExportXML(item)}><FileText size={14} /></button>
          {isAdmin && <>
            <button style={S.iconBtn} title="Editar" onClick={() => onEdit(item)}><Edit size={14} /></button>
            <button style={S.iconBtn} title="Duplicar" onClick={() => onDuplicate(item)}><Copy size={14} /></button>
            <button style={{ ...S.iconBtn, marginLeft: "auto", color: "#8a2f2f" }} title="Excluir" onClick={() => onDelete(item)}><Trash2 size={14} /></button>
          </>}
        </div>
      </div>
    </div>
  );
}

function ControlePage({ allItems, onView, onEdit, onDuplicate, onDelete, onPrint, onExportXML, responsaveis }) {
  const [nome, setNome] = useState(""); const [codigo, setCodigo] = useState("");
  const [tipo, setTipo] = useState(""); const [grupo, setGrupo] = useState("");
  const [dataIni, setDataIni] = useState(""); const [dataFim, setDataFim] = useState("");
  const [resp, setResp] = useState("");

  const filtered = allItems.filter((i) => {
    if (nome && !(i.nome || "").toLowerCase().includes(nome.toLowerCase())) return false;
    if (codigo && !(i.codigo || "").toLowerCase().includes(codigo.toLowerCase())) return false;
    if (tipo && i.tipoAmostra !== tipo) return false;
    if (grupo) { const g = i.tipoAmostra === "mineral" ? i.grupo : i.tipoRocha; if (g !== grupo) return false; }
    if (dataIni && (i.dataCadastro || "") < dataIni) return false;
    if (dataFim && (i.dataCadastro || "") > dataFim) return false;
    if (resp && i.responsavel !== resp) return false;
    return true;
  });

  return (
    <div>
      <PageHeader title="Controle de Cadastros" subtitle="Pesquise, filtre e gerencie todos os registros do acervo" />
      <div style={S.card}>
        <div style={S.filterGrid}>
          <div><label style={S.labelSm}>Nome</label><input style={S.inputSm} value={nome} onChange={(e) => setNome(e.target.value)} placeholder="Buscar por nome…" /></div>
          <div><label style={S.labelSm}>Código</label><input style={S.inputSm} value={codigo} onChange={(e) => setCodigo(e.target.value)} placeholder="Ex: MIN001" /></div>
          <div><label style={S.labelSm}>Tipo</label>
            <select style={S.inputSm} value={tipo} onChange={(e) => { setTipo(e.target.value); setGrupo(""); }}>
              <option value="">Todos</option><option value="mineral">Mineral</option><option value="rocha">Rocha</option>
            </select>
          </div>
          <div><label style={S.labelSm}>Grupo mineralógico / Tipo de rocha</label>
            <select style={S.inputSm} value={grupo} onChange={(e) => setGrupo(e.target.value)}>
              <option value="">Todos</option>
              {(tipo === "rocha" ? TIPOS_ROCHA : tipo === "mineral" ? GRUPOS_MINERALOGICOS : [...GRUPOS_MINERALOGICOS, ...TIPOS_ROCHA]).map((g) => <option key={g} value={g}>{g}</option>)}
            </select>
          </div>
          <div><label style={S.labelSm}>Data inicial</label><input type="date" style={S.inputSm} value={dataIni} onChange={(e) => setDataIni(e.target.value)} /></div>
          <div><label style={S.labelSm}>Data final</label><input type="date" style={S.inputSm} value={dataFim} onChange={(e) => setDataFim(e.target.value)} /></div>
          <div><label style={S.labelSm}>Responsável</label>
            <select style={S.inputSm} value={resp} onChange={(e) => setResp(e.target.value)}>
              <option value="">Todos</option>
              {responsaveis.map((r) => <option key={r} value={r}>{r}</option>)}
            </select>
          </div>
        </div>
      </div>

      <div style={S.card}>
        <div style={S.cardHeadRow}><h3 style={S.cardTitle}>{filtered.length} registro(s) encontrado(s)</h3></div>
        {filtered.length === 0 ? <EmptyState text="Nenhum registro corresponde aos filtros aplicados." /> : (
          <div style={S.tableWrap}>
            <table style={S.table}>
              <thead><tr><th>Código</th><th>Nome</th><th>Tipo</th><th>Cor</th><th>Responsável</th><th>Data</th><th style={{ textAlign: "right" }}>Ações</th></tr></thead>
              <tbody>
                {filtered.map((it) => (
                  <tr key={it.id}>
                    <td style={S.mono}>{it.codigo}</td>
                    <td style={{ fontWeight: 600 }}>{it.nome}</td>
                    <td><TypeBadge item={it} /></td>
                    <td>{it.cor || "—"}</td>
                    <td>{it.responsavel}</td>
                    <td>{fmtDate(it.dataCadastro)}</td>
                    <td>
                      <div style={{ display: "flex", gap: 4, justifyContent: "flex-end" }}>
                        <button style={S.iconBtn} title="Visualizar" onClick={() => onView(it)}><Eye size={13} /></button>
                        <button style={S.iconBtn} title="Editar" onClick={() => onEdit(it)}><Edit size={13} /></button>
                        <button style={S.iconBtn} title="Duplicar" onClick={() => onDuplicate(it)}><Copy size={13} /></button>
                        <button style={S.iconBtn} title="Imprimir" onClick={() => onPrint(it)}><Printer size={13} /></button>
                        <button style={S.iconBtn} title="Exportar XML" onClick={() => onExportXML(it)}><FileText size={13} /></button>
                        <button style={{ ...S.iconBtn, color: "#8a2f2f" }} title="Excluir" onClick={() => onDelete(it)}><Trash2 size={13} /></button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

function RelatoriosPage({ minerais, rochas, exportExcel, exportXMLAll }) {
  const porGrupo = useMemo(() => {
    const m = {};
    minerais.forEach((x) => { m[x.grupo || "Não informado"] = (m[x.grupo || "Não informado"] || 0) + 1; });
    return Object.entries(m).sort((a, b) => b[1] - a[1]);
  }, [minerais]);

  const porResponsavel = useMemo(() => {
    const m = {};
    [...minerais, ...rochas].forEach((x) => { m[x.responsavel || "Não informado"] = (m[x.responsavel || "Não informado"] || 0) + 1; });
    return Object.entries(m).sort((a, b) => b[1] - a[1]);
  }, [minerais, rochas]);

  const flat = (arr, tipo) => arr.map((x) => ({
    Código: x.codigo, Nome: x.nome, Tipo: tipo, Grupo_Tipo: tipo === "Mineral" ? x.grupo : x.tipoRocha,
    Cor: x.cor, Responsável: x.responsavel, Data: fmtDate(x.dataCadastro), Local: tipo === "Mineral" ? x.localArmazenamento : x.localFisico,
  }));

  const printReport = () => window.print();

  return (
    <div>
      <PageHeader title="Relatórios" subtitle="Consolidação do acervo e exportação para PDF, Excel e XML" right={
        <div style={{ display: "flex", gap: 8 }}>
          <button style={S.btnGhostSm} onClick={printReport}><Printer size={15} /> PDF</button>
          <button style={S.btnGhostSm} onClick={() => exportExcel([...flat(minerais, "Mineral"), ...flat(rochas, "Rocha")], "Acervo", `acervo_${todayISO()}.xlsx`)}><FileDown size={15} /> Excel</button>
          <button style={S.btnGhostSm} onClick={exportXMLAll}><FileText size={15} /> XML</button>
        </div>
      } />

      <div style={S.statsGrid}>
        <StatCard icon={Gem} label="Total de minerais" value={minerais.length} accent="#2E7D5B" />
        <StatCard icon={Mountain} label="Total de rochas" value={rochas.length} accent="#C9A227" />
        <StatCard icon={Boxes} label="Total geral" value={minerais.length + rochas.length} accent="#0B3D2E" />
      </div>

      <div style={S.reportGrid}>
        <div style={S.card}>
          <h3 style={S.cardTitle}>Cadastros por grupo mineralógico</h3>
          {porGrupo.length === 0 ? <EmptyState text="Sem dados suficientes." /> : (
            <table style={S.table}><thead><tr><th>Grupo</th><th style={{ textAlign: "right" }}>Qtd.</th></tr></thead>
              <tbody>{porGrupo.map(([g, n]) => <tr key={g}><td>{g}</td><td style={{ textAlign: "right", fontWeight: 700 }}>{n}</td></tr>)}</tbody>
            </table>
          )}
        </div>
        <div style={S.card}>
          <h3 style={S.cardTitle}>Cadastros por responsável</h3>
          {porResponsavel.length === 0 ? <EmptyState text="Sem dados suficientes." /> : (
            <table style={S.table}><thead><tr><th>Responsável</th><th style={{ textAlign: "right" }}>Qtd.</th></tr></thead>
              <tbody>{porResponsavel.map(([r, n]) => <tr key={r}><td>{r}</td><td style={{ textAlign: "right", fontWeight: 700 }}>{n}</td></tr>)}</tbody>
            </table>
          )}
        </div>
      </div>

      <div style={S.card}>
        <h3 style={S.cardTitle}>Quantidade de amostras por tipo</h3>
        <div style={S.barRow}><span style={{ width: 110 }}>Minerais</span><div style={S.barTrack}><div style={{ ...S.barFill, width: `${(minerais.length / Math.max(1, minerais.length + rochas.length)) * 100}%`, background: "#2E7D5B" }} /></div><strong>{minerais.length}</strong></div>
        <div style={S.barRow}><span style={{ width: 110 }}>Rochas</span><div style={S.barTrack}><div style={{ ...S.barFill, width: `${(rochas.length / Math.max(1, minerais.length + rochas.length)) * 100}%`, background: "#C9A227" }} /></div><strong>{rochas.length}</strong></div>
      </div>
    </div>
  );
}

function HistoricoPage({ historico }) {
  return (
    <div>
      <PageHeader title="Histórico de Alterações" subtitle="Registro de ações realizadas no sistema (últimas 500)" />
      <div style={S.card}>
        {historico.length === 0 ? <EmptyState text="Nenhuma alteração registrada ainda." icon={History} /> : (
          <div style={S.tableWrap}>
            <table style={S.table}>
              <thead><tr><th>Data/Hora</th><th>Usuário</th><th>Ação</th><th>Tipo</th><th>Código</th><th>Detalhes</th></tr></thead>
              <tbody>
                {historico.map((h) => (
                  <tr key={h.id}>
                    <td>{new Date(h.data).toLocaleString("pt-BR")}</td>
                    <td>{h.usuario}</td>
                    <td><span style={S.badge}>{h.acao}</span></td>
                    <td>{h.tipo}</td>
                    <td style={S.mono}>{h.codigo}</td>
                    <td style={{ color: "#5b7768" }}>{h.detalhes}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

function UsuariosPage({ usuarios, onSave, onRevoke, authId }) {
  const [form, setForm] = useState(null);
  const pendentes = usuarios.filter((u) => u.status === "Pendente");
  const demais = usuarios.filter((u) => u.status !== "Pendente");
  const statusBadge = (status) => {
    const cfg = status === "Pendente" ? { bg: "#FBF2DA", c: "#7a5f0a" } : status === "Recusado" ? { bg: "#FBEAEA", c: "#8a2f2f" } : { bg: "#E4F2E9", c: "#0B3D2E" };
    return <span style={{ ...S.badge, background: cfg.bg, color: cfg.c }}>{status || "Aprovado"}</span>;
  };
  return (
    <div>
      <PageHeader title="Usuários e Permissões" subtitle="Contas criadas pelos próprios usuários na tela de login (aba 'Criar conta') e aprovadas aqui" />

      <div style={{ ...S.card, borderColor: "#e8dcae", background: "#FFFCF2" }}>
        <p style={{ fontSize: 12.5, color: "#6b5817", margin: 0, lineHeight: 1.7 }}>
          As senhas agora são gerenciadas com segurança pelo Supabase Auth (criptografadas), então não é possível
          cadastrar usuários manualmente por aqui nem ver/alterar a senha de ninguém — cada pessoa cria a própria
          conta pela tela de login. Você pode aprovar, recusar, promover a Administrador ou revogar o acesso.
        </p>
      </div>

      {pendentes.length > 0 && (
        <div style={{ ...S.card, borderColor: "#e8dcae", background: "#FFFCF2" }}>
          <h3 style={S.cardTitle}><TriangleAlert size={15} color="#9a7b0a" style={{ verticalAlign: -2 }} /> {pendentes.length} solicitação(ões) de cadastro aguardando aprovação</h3>
          <div style={S.tableWrap}>
            <table style={S.table}>
              <thead><tr><th>Nome</th><th>E-mail</th><th style={{ textAlign: "right" }}>Ações</th></tr></thead>
              <tbody>
                {pendentes.map((u) => (
                  <tr key={u.id}>
                    <td style={{ fontWeight: 600 }}>{u.nome}</td>
                    <td style={S.mono}>{u.email}</td>
                    <td>
                      <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                        <button style={{ ...S.iconBtn, background: "#E4F2E9", color: "#0B3D2E" }} title="Aprovar" onClick={() => onSave({ ...u, status: "Aprovado" })}><CheckCircle size={14} /></button>
                        <button style={{ ...S.iconBtn, background: "#FBEAEA", color: "#8a2f2f" }} title="Recusar" onClick={() => onSave({ ...u, status: "Recusado" })}><X size={14} /></button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      <div style={S.card}>
        <div style={S.tableWrap}>
          <table style={S.table}>
            <thead><tr><th>Nome</th><th>E-mail</th><th>Perfil</th><th>Status</th><th></th></tr></thead>
            <tbody>
              {demais.map((u) => (
                <tr key={u.id}>
                  <td style={{ fontWeight: 600 }}>{u.nome}</td>
                  <td style={S.mono}>{u.email}</td>
                  <td><span style={{ ...S.badge, background: u.role === "Administrador" ? "#E4F2E9" : "#EFEFEF", color: u.role === "Administrador" ? "#0B3D2E" : "#555" }}>{u.role === "Administrador" && <Shield size={11} />} {u.role}</span></td>
                  <td>{statusBadge(u.status)}</td>
                  <td>
                    <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                      <button style={S.iconBtn} title="Editar perfil/status" onClick={() => setForm(u)}><Edit size={14} /></button>
                      <button style={{ ...S.iconBtn, color: "#8a2f2f" }} title="Revogar acesso" onClick={() => onRevoke(u)} disabled={u.id === authId}><Lock size={14} /></button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {form && (
        <ModalShell title="Editar usuário" onClose={() => setForm(null)}>
          <div style={S.formGrid2}>
            <Field label="Nome completo"><input style={S.input} value={form.nome} onChange={(e) => setForm({ ...form, nome: e.target.value })} /></Field>
            <Field label="E-mail"><input style={S.input} value={form.email} disabled /></Field>
            <Field label="Perfil de acesso">
              <select style={S.input} value={form.role} onChange={(e) => setForm({ ...form, role: e.target.value })}>
                <option>Usuário</option><option>Administrador</option>
              </select>
            </Field>
            <Field label="Status">
              <select style={S.input} value={form.status || "Aprovado"} onChange={(e) => setForm({ ...form, status: e.target.value })}>
                <option>Aprovado</option><option>Pendente</option><option>Recusado</option>
              </select>
            </Field>
          </div>
          <div style={S.modalFooter}>
            <button style={S.btnGhost} onClick={() => setForm(null)}>Cancelar</button>
            <button style={S.btnPrimary} onClick={() => { onSave(form); setForm(null); }}><Save size={16} /> Salvar</button>
          </div>
        </ModalShell>
      )}
    </div>
  );
}

function BackupPage({ onBackup, onRestore, isAdmin, stats }) {
  const fileRef = useRef(null);
  return (
    <div>
      <PageHeader title="Backup do Banco de Dados" subtitle="Exporte ou restaure o acervo completo em formato JSON" />
      <div style={S.reportGrid}>
        <div style={S.card}>
          <h3 style={S.cardTitle}><Download size={16} style={{ verticalAlign: -3 }} /> Exportar backup completo</h3>
          <p style={{ color: "#5b7768", fontSize: 13.5, lineHeight: 1.6 }}>Gera um arquivo com todos os {stats.total} registros, usuários e histórico de alterações. Guarde em local seguro.</p>
          <button style={S.btnPrimary} onClick={onBackup}><Database size={16} /> Gerar backup agora</button>
        </div>
        <div style={S.card}>
          <h3 style={S.cardTitle}><Upload size={16} style={{ verticalAlign: -3 }} /> Restaurar backup</h3>
          <p style={{ color: "#5b7768", fontSize: 13.5, lineHeight: 1.6 }}>{isAdmin ? "Substitui todos os dados atuais pelo conteúdo do arquivo selecionado." : "Apenas administradores podem restaurar backups."}</p>
          <input ref={fileRef} type="file" accept=".json" style={{ display: "none" }} onChange={(e) => e.target.files[0] && onRestore(e.target.files[0])} />
          <button style={S.btnGhost} disabled={!isAdmin} onClick={() => fileRef.current?.click()}><FileUp size={16} /> Selecionar arquivo .json</button>
        </div>
      </div>
      <div style={{ ...S.card, borderColor: "#e8dcae", background: "#FFFCF2" }}>
        <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
          <TriangleAlert size={18} color="#9a7b0a" style={{ flexShrink: 0, marginTop: 2 }} />
          <p style={{ color: "#6b5817", fontSize: 13, lineHeight: 1.7, margin: 0 }}>
            Este ambiente de demonstração armazena os dados em um banco de dados chave-valor compartilhado entre os usuários do sistema. Para produção em escala (10.000+ amostras, múltiplos laboratórios, autenticação segura), recomenda-se migrar para um banco relacional dedicado (PostgreSQL/MySQL) com backend próprio — a estrutura de dados aqui já foi projetada para essa migração direta.
          </p>
        </div>
      </div>
    </div>
  );
}

// ---------- Formulários ----------
function Field({ label, children, span }) {
  return <div style={{ gridColumn: span ? `span ${span}` : undefined }}><label style={S.labelSm}>{label}</label>{children}</div>;
}

function ModalShell({ title, onClose, children, wide }) {
  return (
    <div style={S.modalOverlay} onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div style={{ ...S.modal, maxWidth: wide ? 920 : 620 }}>
        <div style={S.modalHead}>
          <h3 style={{ fontFamily: "Fraunces", fontSize: 19, color: "#0B3D2E", margin: 0 }}>{title}</h3>
          <button style={S.iconBtn} onClick={onClose}><X size={17} /></button>
        </div>
        <div style={S.modalBody}>{children}</div>
      </div>
    </div>
  );
}

function PhotoPicker({ label, value, onChange, multiple }) {
  const inputRef = useRef(null);
  const handleFiles = async (files) => {
    const arr = Array.from(files);
    if (multiple) {
      const news = await Promise.all(arr.map((f) => resizeImage(f)));
      onChange([...(value || []), ...news].slice(0, 6));
    } else {
      const img = await resizeImage(arr[0]);
      onChange(img);
    }
  };
  return (
    <div>
      <label style={S.labelSm}>{label}</label>
      <div style={S.photoDrop} onClick={() => inputRef.current?.click()}>
        <Upload size={16} color="#5b7768" /> <span>Clique para enviar {multiple ? "imagens" : "uma imagem"}</span>
      </div>
      <input ref={inputRef} type="file" accept="image/*" multiple={multiple} style={{ display: "none" }} onChange={(e) => e.target.files.length && handleFiles(e.target.files)} />
      {multiple ? (
        <div style={{ display: "flex", gap: 8, marginTop: 8, flexWrap: "wrap" }}>
          {(value || []).map((v, i) => (
            <div key={i} style={S.thumbWrap}>
              <img src={v} style={S.thumb} />
              <button style={S.thumbRemove} onClick={() => onChange(value.filter((_, idx) => idx !== i))}><X size={11} /></button>
            </div>
          ))}
        </div>
      ) : value ? (
        <div style={{ ...S.thumbWrap, marginTop: 8 }}>
          <img src={value} style={S.thumb} />
          <button style={S.thumbRemove} onClick={() => onChange("")}><X size={11} /></button>
        </div>
      ) : null}
    </div>
  );
}

function MohsSlider({ value, onChange }) {
  return (
    <div>
      <label style={S.labelSm}>Dureza (Escala de Mohs) {value && <span style={{ color: "#0B3D2E", fontWeight: 700 }}>— {value} ({MOHS_REF.find((m) => String(m.d) === String(value))?.m || ""})</span>}</label>
      <div style={S.mohsRow}>
        {MOHS_REF.map((m) => (
          <button key={m.d} type="button" onClick={() => onChange(String(m.d))} title={`${m.d} — ${m.m}`}
            style={{ ...S.mohsFacet, background: Number(value) >= m.d ? "#2E7D5B" : "#E4F2E9", color: Number(value) >= m.d ? "#fff" : "#5b7768" }}>
            {m.d}
          </button>
        ))}
      </div>
    </div>
  );
}

function MineralFormModal({ data, onClose, onSave, auth }) {
  const [f, setF] = useState({ ...data, dataCadastro: data.dataCadastro || todayISO(), responsavel: data.responsavel || auth.nome });
  const set = (k) => (e) => setF({ ...f, [k]: e?.target ? e.target.value : e });
  const submit = () => { if (!f.nome) return; onSave(f); };
  return (
    <ModalShell title={f.id ? `Editar mineral — ${f.codigo}` : "Novo Mineral"} onClose={onClose} wide>
      <div>
        <SectionTitle icon={Gem}>Identificação</SectionTitle>
        <div style={S.formGrid3}>
          <Field label="Código"><input style={S.input} value={f.codigo} onChange={set("codigo")} placeholder="Automático se vazio" /></Field>
          <Field label="Nome do mineral *"><input style={S.input} value={f.nome} onChange={set("nome")} required /></Field>
          <Field label="Nome científico"><input style={S.input} value={f.nomeCientifico} onChange={set("nomeCientifico")} /></Field>
          <Field label="Grupo mineralógico"><select style={S.input} value={f.grupo} onChange={set("grupo")}><option value="">Selecione…</option>{GRUPOS_MINERALOGICOS.map((g) => <option key={g}>{g}</option>)}</select></Field>
          <Field label="Classe química"><select style={S.input} value={f.classeQuimica} onChange={set("classeQuimica")}><option value="">Selecione…</option>{CLASSES_QUIMICAS.map((g) => <option key={g}>{g}</option>)}</select></Field>
        </div>

        <SectionTitle icon={ScanLine}>Características físicas</SectionTitle>
        <div style={S.formGrid3}>
          <Field label="Cor"><input style={S.input} value={f.cor} onChange={set("cor")} /></Field>
          <Field label="Traço"><input style={S.input} value={f.traco} onChange={set("traco")} /></Field>
          <Field label="Brilho"><select style={S.input} value={f.brilho} onChange={set("brilho")}><option value="">Selecione…</option>{BRILHOS.map((g) => <option key={g}>{g}</option>)}</select></Field>
          <Field label="Transparência"><select style={S.input} value={f.transparencia} onChange={set("transparencia")}><option value="">Selecione…</option>{TRANSPARENCIAS.map((g) => <option key={g}>{g}</option>)}</select></Field>
          <Field label="Clivagem"><select style={S.input} value={f.clivagem} onChange={set("clivagem")}><option value="">Selecione…</option>{CLIVAGENS.map((g) => <option key={g}>{g}</option>)}</select></Field>
          <Field label="Fratura"><select style={S.input} value={f.fratura} onChange={set("fratura")}><option value="">Selecione…</option>{FRATURAS.map((g) => <option key={g}>{g}</option>)}</select></Field>
          <Field label="Hábito cristalino"><input style={S.input} value={f.habito} onChange={set("habito")} /></Field>
          <Field label="Sistema cristalino"><select style={S.input} value={f.sistemaCristalino} onChange={set("sistemaCristalino")}><option value="">Selecione…</option>{SISTEMAS_CRISTALINOS.map((g) => <option key={g}>{g}</option>)}</select></Field>
          <Field label="Densidade (g/cm³)"><input style={S.input} type="number" step="0.01" value={f.densidade} onChange={set("densidade")} /></Field>
          <Field label="Magnetismo"><select style={S.input} value={f.magnetismo} onChange={set("magnetismo")}><option>Não</option><option>Sim</option><option>Fraco</option></select></Field>
          <Field label="Fluorescência"><select style={S.input} value={f.fluorescencia} onChange={set("fluorescencia")}><option>Não</option><option>Sim</option></select></Field>
          <Field label="Tenacidade"><select style={S.input} value={f.tenacidade} onChange={set("tenacidade")}><option value="">Selecione…</option>{TENACIDADES.map((g) => <option key={g}>{g}</option>)}</select></Field>
        </div>
        <div style={{ marginTop: 12 }}><MohsSlider value={f.dureza} onChange={(v) => setF({ ...f, dureza: v })} /></div>

        <SectionTitle icon={FileText}>Informações complementares</SectionTitle>
        <div style={S.formGrid2}>
          <Field label="Composição química"><input style={S.input} value={f.composicaoQuimica} onChange={set("composicaoQuimica")} placeholder="Ex: SiO₂" /></Field>
          <Field label="Ambiente de formação"><input style={S.input} value={f.ambienteFormacao} onChange={set("ambienteFormacao")} /></Field>
          <Field label="Principais usos"><input style={S.input} value={f.usos} onChange={set("usos")} /></Field>
          <Field label="Principais ocorrências"><input style={S.input} value={f.ocorrencias} onChange={set("ocorrencias")} /></Field>
          <Field label="Estado de conservação"><select style={S.input} value={f.estadoConservacao} onChange={set("estadoConservacao")}><option value="">Selecione…</option>{ESTADOS_CONSERVACAO.map((g) => <option key={g}>{g}</option>)}</select></Field>
        </div>

        <SectionTitle icon={MapPin}>Controle de armazenamento</SectionTitle>
        <div style={S.formGrid3}>
          <Field label="Local de armazenamento"><input style={S.input} value={f.localArmazenamento} onChange={set("localArmazenamento")} /></Field>
          <Field label="Armário"><input style={S.input} value={f.armario} onChange={set("armario")} /></Field>
          <Field label="Gaveta"><input style={S.input} value={f.gaveta} onChange={set("gaveta")} /></Field>
          <Field label="Prateleira"><input style={S.input} value={f.prateleira} onChange={set("prateleira")} /></Field>
          <Field label="Número da caixa"><input style={S.input} value={f.caixa} onChange={set("caixa")} /></Field>
        </div>

        <SectionTitle icon={ClipboardList}>Registro</SectionTitle>
        <div style={S.formGrid3}>
          <Field label="Data do cadastro"><input type="date" style={S.input} value={f.dataCadastro} onChange={set("dataCadastro")} /></Field>
          <Field label="Responsável pelo cadastro"><input style={S.input} value={f.responsavel} onChange={set("responsavel")} /></Field>
          <Field label="Observações" span={3}><textarea style={{ ...S.input, minHeight: 60 }} value={f.observacoes} onChange={set("observacoes")} /></Field>
        </div>

        <SectionTitle icon={ImageIcon}>Mídia</SectionTitle>
        <div style={S.formGrid2}>
          <PhotoPicker label="Foto principal" value={f.fotoPrincipal} onChange={(v) => setF({ ...f, fotoPrincipal: v })} />
          <PhotoPicker label="Fotos adicionais (até 6)" value={f.fotosAdicionais} onChange={(v) => setF({ ...f, fotosAdicionais: v })} multiple />
        </div>

        <div style={S.modalFooter}>
          <button type="button" style={S.btnGhost} onClick={onClose}>Cancelar</button>
          <button type="button" style={S.btnPrimary} onClick={submit}><Save size={16} /> Salvar mineral</button>
        </div>
      </div>
    </ModalShell>
  );
}

function RochaFormModal({ data, onClose, onSave, auth }) {
  const [f, setF] = useState({ ...data, dataCadastro: data.dataCadastro || todayISO(), responsavel: data.responsavel || auth.nome });
  const set = (k) => (e) => setF({ ...f, [k]: e?.target ? e.target.value : e });
  const submit = () => { if (!f.nome) return; onSave(f); };
  return (
    <ModalShell title={f.id ? `Editar rocha — ${f.codigo}` : "Nova Rocha"} onClose={onClose} wide>
      <div>
        <SectionTitle icon={Mountain}>Identificação</SectionTitle>
        <div style={S.formGrid3}>
          <Field label="Código"><input style={S.input} value={f.codigo} onChange={set("codigo")} placeholder="Automático se vazio" /></Field>
          <Field label="Nome *"><input style={S.input} value={f.nome} onChange={set("nome")} required /></Field>
          <Field label="Tipo"><select style={S.input} value={f.tipoRocha} onChange={set("tipoRocha")}><option value="">Selecione…</option>{TIPOS_ROCHA.map((g) => <option key={g}>{g}</option>)}</select></Field>
        </div>

        <SectionTitle icon={ScanLine}>Características</SectionTitle>
        <div style={S.formGrid3}>
          <Field label="Cor"><input style={S.input} value={f.cor} onChange={set("cor")} /></Field>
          <Field label="Granulação"><select style={S.input} value={f.granulacao} onChange={set("granulacao")}><option value="">Selecione…</option>{GRANULACOES.map((g) => <option key={g}>{g}</option>)}</select></Field>
          <Field label="Textura"><select style={S.input} value={f.textura} onChange={set("textura")}><option value="">Selecione…</option>{TEXTURAS_ROCHA.map((g) => <option key={g}>{g}</option>)}</select></Field>
          <Field label="Estrutura"><input style={S.input} value={f.estrutura} onChange={set("estrutura")} /></Field>
          <Field label="Mineralogia principal"><input style={S.input} value={f.mineralogiaPrincipal} onChange={set("mineralogiaPrincipal")} /></Field>
          <Field label="Mineralogia secundária"><input style={S.input} value={f.mineralogiaSecundaria} onChange={set("mineralogiaSecundaria")} /></Field>
        </div>

        <SectionTitle icon={FileText}>Formação</SectionTitle>
        <div style={S.formGrid3}>
          <Field label="Ambiente de formação"><input style={S.input} value={f.ambienteFormacao} onChange={set("ambienteFormacao")} /></Field>
          <Field label="Idade geológica"><input style={S.input} value={f.idadeGeologica} onChange={set("idadeGeologica")} /></Field>
          <Field label="Origem"><input style={S.input} value={f.origem} onChange={set("origem")} /></Field>
        </div>

        <SectionTitle icon={Boxes}>Aplicações</SectionTitle>
        <div style={S.formGrid3}>
          <Field label="Uso industrial"><input style={S.input} value={f.usoIndustrial} onChange={set("usoIndustrial")} /></Field>
          <Field label="Uso ornamental"><input style={S.input} value={f.usoOrnamental} onChange={set("usoOrnamental")} /></Field>
          <Field label="Uso na construção"><input style={S.input} value={f.usoConstrucao} onChange={set("usoConstrucao")} /></Field>
        </div>

        <SectionTitle icon={MapPin}>Localização de origem</SectionTitle>
        <div style={S.formGrid3}>
          <Field label="Procedência"><input style={S.input} value={f.procedencia} onChange={set("procedencia")} /></Field>
          <Field label="Município"><input style={S.input} value={f.municipio} onChange={set("municipio")} /></Field>
          <Field label="Estado"><select style={S.input} value={f.estado} onChange={set("estado")}><option value="">Selecione…</option>{ESTADOS_BR.map((g) => <option key={g}>{g}</option>)}</select></Field>
          <Field label="País"><input style={S.input} value={f.pais} onChange={set("pais")} /></Field>
        </div>

        <SectionTitle icon={Boxes}>Controle físico</SectionTitle>
        <div style={S.formGrid3}>
          <Field label="Local físico"><input style={S.input} value={f.localFisico} onChange={set("localFisico")} /></Field>
          <Field label="Armário"><input style={S.input} value={f.armario} onChange={set("armario")} /></Field>
          <Field label="Gaveta"><input style={S.input} value={f.gaveta} onChange={set("gaveta")} /></Field>
          <Field label="Prateleira"><input style={S.input} value={f.prateleira} onChange={set("prateleira")} /></Field>
          <Field label="Caixa"><input style={S.input} value={f.caixa} onChange={set("caixa")} /></Field>
        </div>

        <SectionTitle icon={ClipboardList}>Registro</SectionTitle>
        <div style={S.formGrid3}>
          <Field label="Data do cadastro"><input type="date" style={S.input} value={f.dataCadastro} onChange={set("dataCadastro")} /></Field>
          <Field label="Responsável"><input style={S.input} value={f.responsavel} onChange={set("responsavel")} /></Field>
          <Field label="Observações" span={3}><textarea style={{ ...S.input, minHeight: 60 }} value={f.observacoes} onChange={set("observacoes")} /></Field>
        </div>

        <SectionTitle icon={ImageIcon}>Fotos</SectionTitle>
        <PhotoPicker label="Fotos da amostra (até 6)" value={f.fotos} onChange={(v) => setF({ ...f, fotos: v })} multiple />

        <div style={S.modalFooter}>
          <button type="button" style={S.btnGhost} onClick={onClose}>Cancelar</button>
          <button type="button" style={S.btnPrimary} onClick={submit}><Save size={16} /> Salvar rocha</button>
        </div>
      </div>
    </ModalShell>
  );
}

function SectionTitle({ icon: Icon, children }) {
  return <div style={S.sectionTitle}><Icon size={14} /> {children}</div>;
}

function FichaModal({ item, onClose, onEdit, onExportXML }) {
  const isMin = item.tipoAmostra === "mineral";
  const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=180x180&margin=6&data=${encodeURIComponent(item.codigo)}`;
  const fotos = isMin ? [item.fotoPrincipal, ...(item.fotosAdicionais || [])].filter(Boolean) : (item.fotos || []);
  return (
    <div style={S.modalOverlay} onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }} className="printArea">
      <div style={{ ...S.modal, maxWidth: 760 }} id="fichaPrint">
        <div style={S.modalHead} className="noPrint">
          <h3 style={{ fontFamily: "Fraunces", fontSize: 19, color: "#0B3D2E", margin: 0 }}>Ficha da Amostra</h3>
          <div style={{ display: "flex", gap: 6 }}>
            <button style={S.iconBtn} onClick={() => window.print()} title="Imprimir"><Printer size={16} /></button>
            <button style={S.iconBtn} onClick={() => onExportXML(item)} title="Exportar XML"><FileText size={16} /></button>
            <button style={S.iconBtn} onClick={() => onEdit(item)} title="Editar"><Edit size={16} /></button>
            <button style={S.iconBtn} onClick={onClose}><X size={17} /></button>
          </div>
        </div>
        <div style={S.modalBody}>
          <div style={S.fichaHead}>
            <div>
              <div style={{ fontFamily: "Fraunces", fontWeight: 600, fontSize: 24, color: "#0B3D2E" }}>{item.nome}</div>
              <div style={{ color: "#5b7768", fontSize: 13.5 }}>{item.nomeCientifico || (isMin ? item.grupo : item.tipoRocha)}</div>
              <span style={S.mono}>{item.codigo}</span>
            </div>
            <img src={qrUrl} alt="QR Code" style={{ width: 84, height: 84 }} />
          </div>

          {fotos.length > 0 && (
            <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" }}>
              {fotos.map((f, i) => <img key={i} src={f} style={{ width: 100, height: 100, objectFit: "cover", borderRadius: 8, border: "1px solid #dfe8e2" }} />)}
            </div>
          )}

          {isMin ? (
            <>
              <FichaGrupo titulo="Identificação" campos={[["Grupo", item.grupo], ["Classe química", item.classeQuimica]]} />
              <FichaGrupo titulo="Características físicas" campos={[["Cor", item.cor], ["Traço", item.traco], ["Brilho", item.brilho], ["Transparência", item.transparencia], ["Clivagem", item.clivagem], ["Fratura", item.fratura], ["Hábito", item.habito], ["Sistema cristalino", item.sistemaCristalino], ["Dureza (Mohs)", item.dureza], ["Densidade", item.densidade], ["Magnetismo", item.magnetismo], ["Fluorescência", item.fluorescencia], ["Tenacidade", item.tenacidade]]} />
              <FichaGrupo titulo="Complementares" campos={[["Composição química", item.composicaoQuimica], ["Ambiente de formação", item.ambienteFormacao], ["Usos", item.usos], ["Ocorrências", item.ocorrencias], ["Conservação", item.estadoConservacao]]} />
              <FichaGrupo titulo="Armazenamento" campos={[["Local", item.localArmazenamento], ["Armário", item.armario], ["Gaveta", item.gaveta], ["Prateleira", item.prateleira], ["Caixa", item.caixa]]} />
            </>
          ) : (
            <>
              <FichaGrupo titulo="Características" campos={[["Cor", item.cor], ["Granulação", item.granulacao], ["Textura", item.textura], ["Estrutura", item.estrutura], ["Mineralogia principal", item.mineralogiaPrincipal], ["Mineralogia secundária", item.mineralogiaSecundaria]]} />
              <FichaGrupo titulo="Formação" campos={[["Ambiente", item.ambienteFormacao], ["Idade geológica", item.idadeGeologica], ["Origem", item.origem]]} />
              <FichaGrupo titulo="Aplicações" campos={[["Industrial", item.usoIndustrial], ["Ornamental", item.usoOrnamental], ["Construção", item.usoConstrucao]]} />
              <FichaGrupo titulo="Origem" campos={[["Procedência", item.procedencia], ["Município", item.municipio], ["Estado", item.estado], ["País", item.pais]]} />
              <FichaGrupo titulo="Controle físico" campos={[["Local", item.localFisico], ["Armário", item.armario], ["Gaveta", item.gaveta], ["Prateleira", item.prateleira], ["Caixa", item.caixa]]} />
            </>
          )}
          <FichaGrupo titulo="Registro" campos={[["Data do cadastro", fmtDate(item.dataCadastro)], ["Responsável", item.responsavel], ["Observações", item.observacoes]]} />
        </div>
      </div>
    </div>
  );
}

function FichaGrupo({ titulo, campos }) {
  const preenchidos = campos.filter(([, v]) => v);
  if (preenchidos.length === 0) return null;
  return (
    <div style={{ marginBottom: 16 }}>
      <div style={S.fichaSecTitulo}>{titulo}</div>
      <div style={S.fichaGrid}>
        {preenchidos.map(([k, v]) => (
          <div key={k}><div style={S.fichaLabel}>{k}</div><div style={S.fichaValor}>{v}</div></div>
        ))}
      </div>
    </div>
  );
}

// ============================= ESTILOS =============================
const GLOBAL_CSS = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:wght@500;600;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@500&display=swap');
* { box-sizing: border-box; }
body, input, select, textarea, button { font-family: 'Inter', sans-serif; }
input:focus, select:focus, textarea:focus { outline: 2px solid #2E7D5B; outline-offset: 1px; }
button { cursor: pointer; font-family: inherit; }
table { border-collapse: collapse; width: 100%; }
th { text-align: left; font-size: 11.5px; text-transform: uppercase; letter-spacing: 0.4px; color: #5b7768; padding: 10px 12px; border-bottom: 2px solid #E4F2E9; }
td { padding: 10px 12px; font-size: 13.5px; border-bottom: 1px solid #EEF3F0; color: #223028; }
tr:hover td { background: #FAFCF9; }
::-webkit-scrollbar { width: 8px; height: 8px; }
::-webkit-scrollbar-thumb { background: #cfe0d6; border-radius: 8px; }
.crystalSpin { animation: spin 1.6s linear infinite; }
@keyframes spin { to { transform: rotate(360deg); } }
.onlyMobile { display: none; }
@media (max-width: 880px) {
  .onlyMobile { display: inline-flex !important; }
}
@media (max-width: 880px) {
  aside[style] { position: fixed !important; left: -280px; top: 0; bottom: 0; transition: left .2s; }
  aside.mobileNavOpen { left: 0 !important; }
}
@media print {
  body * { visibility: hidden; }
  #fichaPrint, #fichaPrint * { visibility: visible; }
  #fichaPrint { position: absolute; top: 0; left: 0; width: 100%; box-shadow: none !important; }
  .noPrint { display: none !important; }
}
`;

const S = {
  app: { display: "flex", minHeight: "100vh", background: "#F7F9F6", fontFamily: "Inter" },
  bootWrap: { minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#F7F9F6" },
  bootCard: { textAlign: "center" },
  crystal: { width: 46, height: 46, background: "linear-gradient(135deg,#2E7D5B,#0B3D2E)", clipPath: "polygon(50% 0%, 100% 38%, 82% 100%, 18% 100%, 0% 38%)", margin: "0 auto" },

  loginWrap: { minHeight: "100vh", position: "relative", display: "flex", alignItems: "center", justifyContent: "center", background: "#0B3D2E", overflow: "hidden" },
  loginBg: { position: "absolute", inset: 0, background: "radial-gradient(circle at 20% 20%, #1B5E42 0%, #0B3D2E 55%, #072518 100%)" },
  loginCard: { position: "relative", background: "#FFFFFF", borderRadius: 18, padding: "36px 38px", width: 400, maxWidth: "90vw", boxShadow: "0 30px 80px rgba(0,0,0,0.35)" },
  loginBrand: { display: "flex", alignItems: "center", gap: 12 },
  hexIcon: { width: 46, height: 46, background: "linear-gradient(135deg,#2E7D5B,#0B3D2E)", clipPath: "polygon(50% 0%, 100% 25%, 100% 75%, 50% 100%, 0% 75%, 0% 25%)", display: "flex", alignItems: "center", justifyContent: "center" },
  hexIconSm: { width: 34, height: 34, background: "#F7F9F6", clipPath: "polygon(50% 0%, 100% 25%, 100% 75%, 50% 100%, 0% 75%, 0% 25%)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 },
  label: { display: "block", fontSize: 12.5, fontWeight: 600, color: "#0B3D2E", marginBottom: 6, marginTop: 16 },
  labelSm: { display: "block", fontSize: 11.5, fontWeight: 600, color: "#5b7768", marginBottom: 5, textTransform: "uppercase", letterSpacing: 0.3 },
  input: { width: "100%", padding: "10px 12px", borderRadius: 9, border: "1.5px solid #dfe8e2", fontSize: 14, background: "#fff", color: "#1c2a22" },
  inputSm: { width: "100%", padding: "8px 10px", borderRadius: 8, border: "1.5px solid #dfe8e2", fontSize: 13, background: "#fff" },
  loginErr: { display: "flex", alignItems: "center", gap: 6, color: "#8a2f2f", fontSize: 12.5, marginTop: 10, background: "#FBEAEA", padding: "8px 10px", borderRadius: 8 },
  btnPrimary: { display: "flex", alignItems: "center", justifyContent: "center", gap: 8, width: "100%", marginTop: 22, background: "#0B3D2E", color: "#fff", border: "none", padding: "12px 16px", borderRadius: 10, fontWeight: 600, fontSize: 14.5 },
  btnPrimarySm: { display: "flex", alignItems: "center", gap: 7, background: "#0B3D2E", color: "#fff", border: "none", padding: "9px 15px", borderRadius: 9, fontWeight: 600, fontSize: 13 },
  btnGhost: { display: "flex", alignItems: "center", justifyContent: "center", gap: 8, background: "#fff", color: "#0B3D2E", border: "1.5px solid #dfe8e2", padding: "10px 16px", borderRadius: 9, fontWeight: 600, fontSize: 13.5 },
  btnGhostSm: { display: "flex", alignItems: "center", gap: 6, background: "#fff", color: "#0B3D2E", border: "1.5px solid #dfe8e2", padding: "8px 13px", borderRadius: 8, fontWeight: 600, fontSize: 12.5 },
  loginHint: { marginTop: 22, fontSize: 11.5, color: "#5b7768", background: "#F2F7F4", padding: "10px 12px", borderRadius: 9, lineHeight: 1.6 },
  demoBtn: { display: "inline-flex", alignItems: "center", gap: 6, background: "#fff", border: "1px solid #cfe0d6", color: "#0B3D2E", padding: "6px 10px", borderRadius: 7, fontSize: 11.5, fontWeight: 600 },
  authTabs: { display: "flex", gap: 4, marginTop: 24, background: "#F2F7F4", borderRadius: 10, padding: 4 },
  authTab: { flex: 1, background: "transparent", border: "none", padding: "9px 0", borderRadius: 7, fontSize: 13, fontWeight: 600, color: "#5b7768" },
  authTabActive: { background: "#fff", color: "#0B3D2E", boxShadow: "0 2px 6px rgba(11,61,46,0.12)" },
  regOkBox: { textAlign: "center", padding: "24px 10px 8px" },
  forgotLink: { display: "block", width: "100%", textAlign: "center", background: "transparent", border: "none", color: "#5b7768", fontSize: 12.5, fontWeight: 600, marginTop: 12, textDecoration: "underline" },

  sidebar: { background: "#0B3D2E", display: "flex", flexDirection: "column", transition: "width .18s", position: "sticky", top: 0, height: "100vh", flexShrink: 0, zIndex: 20 },
  sidebarTop: { display: "flex", alignItems: "center", gap: 10, padding: "20px 18px", borderBottom: "1px solid rgba(255,255,255,0.08)" },
  nav: { display: "flex", flexDirection: "column", gap: 3, padding: "14px 10px", flex: 1, overflowY: "auto" },
  navItem: { display: "flex", alignItems: "center", gap: 11, background: "transparent", border: "none", color: "#cfe0d6", padding: "10px 12px", borderRadius: 9, fontSize: 13.5, fontWeight: 500, textAlign: "left" },
  navItemActive: { background: "rgba(255,255,255,0.12)", color: "#fff" },
  navBadge: { background: "#C9A227", color: "#fff", fontSize: 10.5, fontWeight: 700, borderRadius: 999, padding: "1px 7px", minWidth: 18, textAlign: "center" },
  sidebarBottom: { padding: 12, borderTop: "1px solid rgba(255,255,255,0.08)", display: "flex", gap: 8 },
  sidebarCollapse: { background: "rgba(255,255,255,0.08)", border: "none", color: "#cfe0d6", borderRadius: 8, width: 34, height: 34, display: "flex", alignItems: "center", justifyContent: "center" },
  logoutBtn: { flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 7, background: "rgba(255,255,255,0.08)", border: "none", color: "#f5d5d5", borderRadius: 8, padding: "8px 10px", fontSize: 13, fontWeight: 600 },
  overlay: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", zIndex: 15 },

  main: { flex: 1, minWidth: 0, display: "flex", flexDirection: "column" },
  topbar: { display: "flex", alignItems: "center", gap: 16, padding: "14px 26px", background: "#fff", borderBottom: "1px solid #E7EEE9", position: "sticky", top: 0, zIndex: 10 },
  menuBtn: { background: "#F2F7F4", border: "none", borderRadius: 8, width: 36, height: 36, alignItems: "center", justifyContent: "center", color: "#0B3D2E" },
  searchBox: { flex: 1, position: "relative", display: "flex", alignItems: "center", gap: 8, background: "#F2F7F4", borderRadius: 10, padding: "9px 14px", maxWidth: 640 },
  searchInput: { flex: 1, border: "none", background: "transparent", fontSize: 13.5, outline: "none" },
  searchResults: { position: "absolute", top: "110%", left: 0, right: 0, background: "#fff", borderRadius: 10, boxShadow: "0 14px 40px rgba(11,61,46,0.18)", border: "1px solid #E7EEE9", maxHeight: 320, overflowY: "auto", zIndex: 30 },
  searchResultRow: { display: "flex", alignItems: "center", gap: 10, padding: "10px 14px", borderBottom: "1px solid #F2F7F4", fontSize: 13 },
  searchEmpty: { padding: 16, fontSize: 13, color: "#5b7768" },
  userChip: { display: "flex", alignItems: "center", gap: 10 },
  userAvatar: { width: 34, height: 34, borderRadius: 9, background: "#0B3D2E", color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 14 },

  content: { padding: "26px 30px 60px", flex: 1 },
  pageHeader: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 22, flexWrap: "wrap", gap: 12 },
  pageTitle: { fontFamily: "Fraunces", fontSize: 26, color: "#0B3D2E", margin: 0, fontWeight: 600 },
  pageSubtitle: { color: "#5b7768", fontSize: 13.5, marginTop: 4 },

  statsGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px,1fr))", gap: 14, marginBottom: 20 },
  statCard: { background: "#fff", border: "1px solid #E7EEE9", borderRadius: 14, padding: "16px 18px", display: "flex", alignItems: "center", gap: 13 },
  statIcon: { width: 44, height: 44, borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 },
  statValue: { fontFamily: "Fraunces", fontSize: 24, fontWeight: 700, color: "#0B3D2E", lineHeight: 1 },
  statLabel: { fontSize: 12, color: "#5b7768", marginTop: 4 },
  statSub: { fontSize: 10.5, color: "#a9bdb2" },

  quickGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(210px,1fr))", gap: 12, marginBottom: 22 },
  quickCard: { display: "flex", alignItems: "center", gap: 12, background: "#fff", border: "1px solid #E7EEE9", borderRadius: 12, padding: "14px 16px", fontSize: 13.5, fontWeight: 600, color: "#1c2a22" },
  quickIcon: { width: 34, height: 34, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 },

  card: { background: "#fff", border: "1px solid #E7EEE9", borderRadius: 14, padding: 20, marginBottom: 18 },
  cardHeadRow: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 },
  cardTitle: { fontFamily: "Fraunces", fontSize: 16.5, color: "#0B3D2E", margin: "0 0 12px", fontWeight: 600 },
  tableWrap: { overflowX: "auto" },
  table: { fontSize: 13.5 },
  mono: { fontFamily: "JetBrains Mono", fontSize: 12, color: "#0B3D2E", fontWeight: 600 },
  badge: { display: "inline-flex", alignItems: "center", gap: 4, background: "#F2F7F4", color: "#0B3D2E", padding: "3px 9px", borderRadius: 999, fontSize: 11.5, fontWeight: 600 },
  iconBtn: { background: "#F2F7F4", border: "none", borderRadius: 7, width: 30, height: 30, display: "inline-flex", alignItems: "center", justifyContent: "center", color: "#0B3D2E" },
  emptyState: { textAlign: "center", padding: "40px 20px", color: "#7c9186", fontSize: 13.5, display: "flex", flexDirection: "column", alignItems: "center", gap: 10 },

  filterRow: { marginBottom: 16 },
  searchBoxSm: { display: "flex", alignItems: "center", gap: 8, background: "#fff", border: "1.5px solid #E7EEE9", borderRadius: 10, padding: "9px 13px", maxWidth: 420 },
  searchInputSm: { flex: 1, border: "none", outline: "none", fontSize: 13 },
  filterGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px,1fr))", gap: 12 },

  cardGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(230px,1fr))", gap: 16 },
  sampleCard: { background: "#fff", border: "1px solid #E7EEE9", borderRadius: 14, overflow: "hidden" },
  sampleImg: { position: "relative", height: 130, background: "#EEF3F0" },
  sampleImgTag: { width: "100%", height: "100%", objectFit: "cover" },
  sampleImgPh: { width: "100%", height: "100%", display: "flex", alignItems: "center", justifyContent: "center" },
  codeChip: { position: "absolute", bottom: 8, left: 8, background: "#fff", border: "1.5px solid", fontFamily: "JetBrains Mono", fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 999, color: "#0B3D2E" },

  reportGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(300px,1fr))", gap: 16, marginBottom: 4 },
  barRow: { display: "flex", alignItems: "center", gap: 12, marginBottom: 10, fontSize: 13 },
  barTrack: { flex: 1, height: 10, background: "#EEF3F0", borderRadius: 6, overflow: "hidden" },
  barFill: { height: "100%", borderRadius: 6 },

  modalOverlay: { position: "fixed", inset: 0, background: "rgba(11,61,46,0.35)", display: "flex", alignItems: "flex-start", justifyContent: "center", padding: "30px 16px", overflowY: "auto", zIndex: 50 },
  modal: { background: "#fff", borderRadius: 16, width: "100%", boxShadow: "0 30px 80px rgba(0,0,0,0.3)" },
  modalHead: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "18px 22px", borderBottom: "1px solid #EEF3F0" },
  modalBody: { padding: "20px 22px 24px" },
  modalFooter: { display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 22, paddingTop: 16, borderTop: "1px solid #EEF3F0" },
  formGrid2: { display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(220px,1fr))", gap: 14 },
  formGrid3: { display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(170px,1fr))", gap: 14 },
  sectionTitle: { display: "flex", alignItems: "center", gap: 7, fontSize: 11.5, fontWeight: 700, color: "#2E7D5B", textTransform: "uppercase", letterSpacing: 0.5, margin: "22px 0 12px", paddingBottom: 8, borderBottom: "1px solid #E4F2E9" },

  photoDrop: { display: "flex", alignItems: "center", gap: 8, border: "1.5px dashed #cfe0d6", borderRadius: 10, padding: "12px 14px", fontSize: 12.5, color: "#5b7768", background: "#FAFCF9" },
  thumbWrap: { position: "relative", width: 60, height: 60 },
  thumb: { width: 60, height: 60, objectFit: "cover", borderRadius: 8, border: "1px solid #dfe8e2" },
  thumbRemove: { position: "absolute", top: -6, right: -6, background: "#8a2f2f", color: "#fff", border: "2px solid #fff", borderRadius: "50%", width: 20, height: 20, display: "flex", alignItems: "center", justifyContent: "center" },

  mohsRow: { display: "flex", gap: 4, marginTop: 4 },
  mohsFacet: { flex: 1, height: 30, border: "none", borderRadius: 4, fontSize: 11.5, fontWeight: 700, clipPath: "polygon(15% 0,100% 0,85% 100%,0 100%)" },

  fichaHead: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 18, paddingBottom: 16, borderBottom: "2px solid #E4F2E9" },
  fichaSecTitulo: { fontSize: 11, fontWeight: 700, color: "#2E7D5B", textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 8 },
  fichaGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(150px,1fr))", gap: 10 },
  fichaLabel: { fontSize: 10.5, color: "#a9bdb2", textTransform: "uppercase", letterSpacing: 0.3 },
  fichaValor: { fontSize: 13.5, color: "#1c2a22", fontWeight: 600, marginTop: 2 },

  toast: { position: "fixed", bottom: 24, right: 24, color: "#fff", padding: "12px 18px", borderRadius: 10, display: "flex", alignItems: "center", gap: 9, fontSize: 13.5, fontWeight: 600, boxShadow: "0 14px 34px rgba(0,0,0,0.25)", zIndex: 100 },
  storageWarn: { position: "fixed", top: 0, left: 0, right: 0, background: "#9a7b0a", color: "#fff", padding: "8px 16px", fontSize: 12.5, display: "flex", alignItems: "center", gap: 8, justifyContent: "center", zIndex: 200 },
};
