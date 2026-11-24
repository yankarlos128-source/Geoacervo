import { useState, useEffect, useCallback, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./lib/supabaseClient";
import {
  Gem, Mountain, LayoutDashboard, Search, Plus, Filter, Printer, Download,
  FileText, Trash2, Edit, Copy, Eye, QrCode, Users, History, Database,
  LogOut, Menu, Moon, Sun, X, Save, Upload, ChevronRight, ChevronLeft, AlertCircle,
  CheckCircle, BarChart3, Package, Calendar, MapPin, Lock, Shield, FileDown,
  FileUp, RefreshCw, Image as ImageIcon, ClipboardList, UserCog, Boxes,
  ArrowUpRight, TriangleAlert, ScanLine, FolderOpen
} from "lucide-react";

// ============================= CONSTANTES =============================
const GRUPOS_MINERALOGICOS = ["Silicatos", "Óxidos", "Sulfetos", "Carbonatos", "Sulfatos", "Haletos", "Fosfatos", "Elementos Nativos", "Boratos", "Outros"];
const CLASSES_QUIMICAS = ["Tectossilicatos", "Filossilicatos", "Inossilicatos", "Ciclossilicatos", "Nesossilicatos", "Sorossilicatos", "Óxidos Simples", "Óxidos Múltiplos", "Sulfetos Simples", "Carbonatos Anidros", "Outra"];
const SISTEMAS_CRISTALINOS = ["Cúbico", "Tetragonal", "Ortorrômbico/Rômbico", "Hexagonal", "Trigonal/Romboédrico", "Monoclínico", "Triclínico", "Amorfo"];
const BRILHOS = ["Metálico", "Não metálico"];
const BRILHOS_NAO_METALICOS = ["Vítreo", "Adamantino", "Resinoso", "Gorduroso", "Sedoso", "Nacarado", "Fosco", "Terroso"];
const REACAO_ACIDO = ["Não reage", "Reage fracamente", "Reage fortemente (efervescência)"];
const TRANSPARENCIAS = ["Transparente", "Translúcido", "Opaco"];
const CLIVAGENS = ["Perfeita", "Boa", "Regular", "Imperfeita", "Ausente"];
const FRATURAS = ["Conchoidal", "Irregular", "Fibrosa", "Ganchosa", "Terrosa"];
const TENACIDADES = ["Frágil", "Maleável", "Séctil", "Dúctil", "Flexível", "Elástico"];
const ESTADOS_CONSERVACAO = ["Excelente", "Bom", "Regular", "Danificado"];
const TIPOS_ROCHA = ["Ígnea", "Sedimentar", "Metamórfica"];
const GRANULACOES = ["Afanítica (fina)", "Fanerítica (grossa)", "Vítrea", "Porfirítica", "Criptocristalina"];
const TEXTURAS_ROCHA = ["Granular", "Foliada", "Bandada", "Maciça", "Vesicular", "Clástica"];
const ESTADOS_BR = ["AC","AL","AP","AM","BA","CE","DF","ES","GO","MA","MT","MS","MG","PA","PB","PR","PE","PI","RJ","RN","RS","RO","RR","SC","SP","SE","TO"];

// Campos específicos por tipo de rocha (cada tipo se analisa de um jeito diferente)
// Ígnea — campos seguindo a Ficha de Descrição de Rochas Ígneas (petrografia)
const ESTRUTURA_VULC_PEQUENA = ["Vesicular", "Amigdaloidal", "Escoriácea", "Perlítica", "Compacta", "Esferulítica", "Fluidal"];
const ESTRUTURA_VULC_GRANDE = ["Cordada", "Almofadada", "Colunar"];
const ESTRUTURA_PLUT_PEQUENA = ["Compacta", "Bandada", "Schlieren"];
const ESTRUTURA_PLUT_MEDIA = ["Enclaves", "Xenólitos"];
const GRAU_CRISTALINIDADE = ["Holocristalina", "Hipocristalina", "Hipovítrea", "Holovítrea"];
const GRAU_VISIBILIDADE = ["Fanerítica", "Afanítica"];
const VISIBILIDADE_TIPOS = { "Fanerítica": ["Equigranular", "Inequigranular", "Megaporfirítica", "Porfirítica"], "Afanítica": ["Vitrofírica"] };
const GRANULACAO_IGNEA = ["Gigantes (>10 cm)", "Muito grossa (3 a 10 cm)", "Grossa (1 a 3 cm)", "Média (1 a 10 mm)", "Fina (0,1 a 1 mm)", "Densa (0,0009 a 0,1 mm)", "Vítrea (N/A)"];
const FORMA_GEOMETRICA_CRISTAIS = ["Idiomórfico", "Hipidiomórfico", "Xenomórfico"];
const AMBIENTE_FORMACAO_IGNEA_OPCOES = ["Vulcânica — resfriamento rápido (cristais pequenos)", "Plutônica — resfriamento moderado (cristais médios)", "Plutônica — resfriamento lento (cristais grandes)"];
const INDICE_COR_IGNEA = ["Leucocrática (M 0-33)", "Mesocrática (M 34-66)", "Melanocrática (M 67-90)"];
const SATURACAO_SILICA = ["Ácidas (teor de sílica maior que 66%)", "Intermediárias (teor de sílica entre 66 e 52%)", "Básicas (teor de sílica entre 52 e 45%)", "Ultrabásicas (teor de sílica menor que 45%)"];
// Sedimentar — campos seguindo o "Roteiro para Prática de Descrição de Rochas Sedimentares"
const GRANULOMETRIA_SED_ANTIGA = ["Conglomerado ou brecha (seixos / grãos maiores)", "Arenito (areia)", "Argilito ou siltito (grãos muito finos)"];
const MATERIAL_FORMACAO_SED_ANTIGA = ["Clástica (fragmentos de rochas ou minerais)", "Química (maciça, sem grãos visíveis)", "Biogênica (fragmentos de conchas ou restos de animais)"];
const SELECIONAMENTO_SED_ANTIGA = ["Bem selecionada (grãos de tamanho igual)", "Moderadamente selecionada", "Mal selecionada (grãos de tamanhos diferentes)"];
const ARREDONDAMENTO_SED = ["Muito angulosos", "Angulosos", "Subangulosos", "Subarredondados", "Arredondados", "Bem arredondados"];
const COR_ORIGEM_SED = ["Primária (cor original)", "Secundária (alterada na superfície)"];
const GRANULOMETRIA_SED = ["Matacão / bloco (> 256 mm)", "Calhau (64 a 256 mm)", "Seixo (4 a 64 mm)", "Grânulo (2 a 4 mm)", "Areia muito grossa (1 a 2 mm)", "Areia grossa (0,5 a 1 mm)", "Areia média (0,25 a 0,5 mm)", "Areia fina (0,125 a 0,25 mm)", "Areia muito fina (0,0625 a 0,125 mm)", "Silte (0,004 a 0,0625 mm)", "Argila (< 0,004 mm)"];
const MATERIAL_FORMACAO_SED = [...MATERIAL_FORMACAO_SED_ANTIGA, "Orgânica (restos vegetais — carvão, turfa)"];
const SELECIONAMENTO_SED = ["Muito bem selecionada", ...SELECIONAMENTO_SED_ANTIGA.slice(0, 1), SELECIONAMENTO_SED_ANTIGA[1], SELECIONAMENTO_SED_ANTIGA[2], "Muito mal selecionada (de argila a seixos misturados)"];
const REACAO_HCL_ROCHA = ["Nenhuma", "Fraca (só no pó ou com gota prolongada)", "Forte (efervescência na rocha inteira)"];
const SED_COESAO = ["Friável (esfarela com os dedos)", "Pouco coesa (desagrega com facilidade)", "Coesa", "Muito coesa (dura, resiste ao martelo)"];
const SED_POROSIDADE = ["Alta (absorve a água rápido)", "Média", "Baixa", "Praticamente nula"];
const SED_ESFERICIDADE = ["Baixa (grãos alongados ou achatados)", "Média", "Alta (grãos quase esféricos)"];
const SED_ARCABOUCO = ["Suportado pelos grãos (os grãos se tocam)", "Suportado pela matriz (grãos “boiando” na lama ou areia fina)", "Não se aplica"];
const SED_CIMENTO = ["Sem cimento visível", "Silicoso (sílica — muito duro)", "Carbonático (calcita — reage ao HCl)", "Ferruginoso (óxido de ferro — avermelhado/amarronzado)", "Argiloso (argila)", "Outro"];
const SED_TEXTURA_NAO_CLASTICA = ["Cristalina (cristais de calcita, dolomita, halita, gipsita…)", "Oolítica (esférulas pequenas, tipo “ovinhos”)", "Microcristalina / criptocristalina (maciça, aspecto de porcelana ou vidro fosco)", "Bioclástica (fragmentos de conchas e esqueletos)", "Coquina (conchas inteiras cimentadas)", "Fibrosa / laminada (restos vegetais)", "Porosa / esponjosa"];
const SED_ACAMAMENTO = ["Maciço (sem estrutura interna visível)", "Laminação plano-paralela (lâminas finas e retas)", "Estratificação plano-paralela (camadas retas)", "Estratificação cruzada (camadas inclinadas cortando outras)", "Gradação normal (grãos mais grossos na base)", "Gradação inversa (grãos mais grossos no topo)", "Laminação ondulada (ondulações / ripples)", "Laminação convoluta (deformada)", "Alternância rítmica (ritmito / varvito)"];
const SED_ESPESSURA = ["Laminada (< 1 cm)", "Muito fina (1 a 3 cm)", "Fina (3 a 10 cm)", "Média (10 a 30 cm)", "Espessa (30 a 100 cm)", "Muito espessa (> 100 cm)"];
const SED_ESTR_SUPERFICIE = ["Nenhuma", "Marcas de onda (ripple marks)", "Gretas de ressecamento (mud cracks)", "Marcas de corrente na base da camada", "Pegadas e escavações de organismos (icnofósseis)", "Marcas de gotas de chuva", "Concreções / nódulos", "Estilólitos"];
const SED_FOSSEIS = ["Ausentes", "Conchas e fragmentos (bioclastos)", "Restos vegetais (folhas, troncos)", "Ossos / dentes (vertebrados)", "Icnofósseis (pegadas, tubos, escavações)", "Microfósseis (só em lupa/microscópio)"];
const SED_CLASSIFICACAO = ["Conglomerado", "Brecha", "Arenito", "Siltito", "Argilito", "Folhelho", "Calcário", "Dolomito", "Evaporito (halita, gipsita, anidrita)", "Chert / sílex", "Carvão", "Formação ferrífera / ironstone", "Fosforito", "Outro"];
const SED_AMBIENTE_DEP = ["Fluvial (rio)", "Lacustre (lago)", "Eólico (deserto / dunas)", "Glacial", "Deltaico", "Costeiro / praial", "Marinho raso (plataforma)", "Marinho profundo", "Lagunar / evaporítico", "Recifal (carbonático)", "Pantanoso / turfeira", "Não identificado"];
function sugerirNomeSed(f) {
  const g = f.sedGranulometria || "", mat = f.sedMaterialFormacao || "", arr = f.sedArredondamento || "";
  const hcl = (f.rochaReacaoHCl || "").startsWith("Forte");
  if (mat.startsWith("Orgânica")) return "Carvão";
  if (mat.startsWith("Química") || mat.startsWith("Biogênica")) {
    if ((f.sedTexturaNaoClastica || "").startsWith("Coquina")) return "Coquina (calcário de conchas)";
    if (hcl) return mat.startsWith("Biogênica") ? "Calcário bioclástico" : "Calcário";
    if ((f.sedTexturaNaoClastica || "").startsWith("Microcristalina") && !hcl) return "Chert / sílex";
    return "";
  }
  if (/^(Matacão|Calhau|Seixo|Grânulo)/.test(g) || g.startsWith("Conglomerado")) {
    if (/Angulosos|Muito angulosos|Subangulosos/.test(arr)) return "Brecha";
    return arr ? "Conglomerado" : "Conglomerado ou brecha";
  }
  if (/^Areia/.test(g) || g.startsWith("Arenito")) return "Arenito";
  if (g.startsWith("Silte")) return "Siltito";
  if (g.startsWith("Argila") || g.startsWith("Argilito")) return /Laminação|plano-paralela/.test(f.sedAcamamento || "") ? "Folhelho (se físsil) ou argilito" : "Argilito";
  return "";
}
// Metamórfica — campos seguindo a "Ficha de Descrição – Rochas Metamórficas"
const MET_GRANULACAO = ["Muito finos (não vejo os grãos a olho nu, tipo ardósia/filito)", "Finos (vejo alguns brilhinhos, ainda pequenos)", "Médios (distingo vários minerais separados)", "Grossos (cristais grandes, bem destacados)"];
const MET_TEXTURA = ["Granoblástica (monte de grãos sem direção)", "Lepidoblástica (plaquinhas/lâminas brilhantes alinhadas)", "Nematoblástica (cristais alongados, tipo agulhas, na mesma direção)", "Porfiroblástica (alguns cristais bem maiores que o resto)", "Cataclástica/milonítica (aspecto de rocha triturada, estirada em um plano)", "Granolepidoblástica (grãos e lamelas juntos)", "Hornfélsica (grãos entrelaçados, sem direção — metamorfismo de contato)"];
const MET_ASPECTO = ["Maciça (tudo misturado sem direção)", "Placas fininhas (clivagem ardosiana)", "Folhas maiores (xistosidade)", "Faixas claras e escuras (bandamento gnáissico)", "Superfície ondulada e sedosa (foliação filítica)", "Minerais estirados com “olhos” (foliação milonítica)"];
const MET_BRILHO = ["Fosco (terroso)", "Sedoso (brilho de seda — típico de filito)", "Micáceo (brilho de lamelas de mica)", "Vítreo (quartzo, granada)", "Gorduroso / graxo (talco, serpentina)", "Metálico (grafita, magnetita)"];
const MET_FORMA_CRISTAIS = ["Idioblásticos (faces cristalinas bem formadas)", "Hipidioblásticos (faces parcialmente formadas)", "Xenoblásticos (sem faces próprias)"];
const MET_PORFIROBLASTOS = ["Ausentes", "Granada", "Estaurolita", "Cianita", "Andaluzita", "Feldspato (estrutura augen)", "Biotita / muscovita", "Quartzo", "Pirita / magnetita (cubos ou octaedros)", "Outro"];
const MET_ESTRUTURAS_ASSOC = ["Nenhuma", "Dobras", "Crenulação (micro-dobras)", "Boudinagem", "Veios / leitos de quartzo", "Lineação mineral", "Estrutura augen (“olhos”)", "Porfiroblastos rotacionados (“bola de neve”)", "Fraturas ou falhas preenchidas"];
const TIPO_METAMORFISMO = ["Regional (grandes áreas — pressão e temperatura)", "De contato / térmico (perto de intrusões)", "Dinâmico / cataclástico (falhas e cisalhamento)", "Hidrotermal / metassomático", "De soterramento", "De impacto"];
const MET_FACIES = ["Zeólita", "Prehnita-pumpellyíta", "Xisto verde", "Anfibolito", "Granulito", "Xisto azul", "Eclogito", "Hornfels (contato)", "Não determinada"];
function sugerirNomeMet(f) {
  const a = f.metAspecto || "", t = f.metTextura || "", p = f.rochaOriginal || "";
  const n = (v) => num(v);
  if (/milonít|Cataclást/i.test(t) || /milonítica/.test(a)) return "Milonito / cataclasito";
  if (/Faixas|gnáissico/.test(a)) return p.startsWith("Ígnea") ? "Gnaisse (ortognaisse)" : p.startsWith("Sedimentar") ? "Gnaisse (paragnaisse)" : "Gnaisse";
  if (/Placas fininhas|ardosiana/.test(a)) return "Ardósia";
  if (/filítica/.test(a)) return "Filito";
  if (/Folhas maiores|xistosidade/.test(a)) return "Xisto";
  if (n(f.metCarbonato) >= 50 || (p.startsWith("Carbonática") && !a.startsWith("Folhas"))) return "Mármore";
  if (n(f.metQuartzo) >= 75) return "Quartzito";
  if (n(f.metAnfibolio) >= 50) return "Anfibolito";
  if (/Hornfél/.test(t)) return "Hornfels";
  if (a.startsWith("Maciça") && p.startsWith("Sedimentar")) return "Quartzito (se rico em quartzo)";
  return "";
}
const GRAU_METAMORFICO = ["Baixo grau (lembra muito o protólito, ex.: ardósia, filito)", "Grau intermediário (minerais maiores e orientações fortes, xisto)", "Alto grau (bandamento, minerais grandes, deformação: gnaisse, migmatito)"];
const MET_PROTOLITO = ["Sedimentar (ex.: argilito → ardósia/filito/xisto)", "Ígnea (ex.: granito → gnaisse)", "Carbonática (calcário → mármore)", "Não foi possível identificar a olho nu"];

const MOHS_REF = [
  { d: 1, m: "Talco" }, { d: 2, m: "Gipsita" }, { d: 3, m: "Calcita" }, { d: 4, m: "Fluorita" },
  { d: 5, m: "Apatita" }, { d: 6, m: "Ortoclásio" }, { d: 7, m: "Quartzo" }, { d: 8, m: "Topázio" },
  { d: 9, m: "Coríndon" }, { d: 10, m: "Diamante" }
];

const emptyMineral = () => ({
  id: "", codigo: "", nome: "", nomeCientifico: "", grupo: "", classeQuimica: "",
  cor: "", traco: "", brilho: "", brilhoDetalhe: "", transparencia: "", clivagem: "", planosClivagem: "", reacaoAcido: "", fratura: "",
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
  // Específicos — Ígnea (segue a Ficha de Descrição de Rochas Ígneas)
  coordenadaX: "", coordenadaY: "", tipoAfloramento: "",
  corAlteracao: "",
  estVulcPequena: "", estVulcGrande: "", estPlutPequena: "", estPlutMedia: "",
  estruturaVulcanica: "", estruturaPlutonica: "", // antigos (só para registros já salvos)
  felFeldspatoK: "", felPlagioclasio: "", felQuartzo: "", felMuscovita: "", felOutroNome: "", felOutroPct: "", felFeldspatoide: "",
  mafTurmalina: "", mafAnfibolio: "", mafBiotita: "", mafMagnetita: "", mafOutroNome: "", mafOutroPct: "",
  grauCristalinidade: "", tamanhoCristaisRelativo: "", tamanhoCristaisAbsoluto: "", grauVisibilidade: "", visibilidadeTipo: "", formaGeometricaCristais: "",
  ambienteFormacaoIgnea: "", indiceCorIgnea: "", saturacaoSilica: "", nomenclaturaIgnea: "",
  // Específicos — Sedimentar (segue o roteiro de descrição de rochas sedimentares)
  sedGranulometria: "", sedMaterialFormacao: "", sedSelecionamento: "", sedArredondamento: "",
  sedCorOrigem: "", sedPossuiEstrutura: "", sedEstruturaDescricao: "", sedEsboco: "",
  sedCor: "", sedCoesao: "", sedPorosidade: "", rochaReacaoHCl: "", sedTexturaNaoClastica: "", sedEsfericidade: "", sedArcabouco: "", sedCimento: "",
  sedAcamamento: "", sedEspessura: "", sedEstrSuperficie: "", sedFosseis: "",
  sedQuartzo: "", sedFeldspato: "", sedLiticos: "", sedMicas: "", sedArgilominerais: "", sedCalcita: "", sedDolomita: "", sedOxidosFe: "", sedMateriaOrganica: "", sedOutroNome: "", sedOutroPct: "",
  sedClassificacao: "", sedAmbienteDep: "", sedNomenclatura: "",
  // Específicos — Metamórfica
  metGranulacao: "", metTextura: "", metAspecto: "", grauMetamorfico: "", rochaOriginal: "", mineraisIndice: "",
  tipoMetamorfismo: "", foliacao: "", // foliacao: campo antigo (só para registros já salvos)
  metCor: "", metBrilho: "", metForma: "", metPorfiroblastos: "", metEstrAssoc: "",
  metQuartzo: "", metFeldspato: "", metCarbonato: "", metMuscovita: "", metBiotita: "", metClorita: "", metAnfibolio: "", metPiroxenio: "", metGranada: "", metTalcoSerp: "", metGrafita: "", metAluminossilicato: "", metOutroNome: "", metOutroPct: "",
  metFacies: "", metNomenclatura: "",
  ambienteFormacao: "", idadeGeologica: "", origem: "",
  usoIndustrial: "", usoOrnamental: "", usoConstrucao: "",
  procedencia: "", municipio: "", estado: "", pais: "Brasil",
  localFisico: "", armario: "", gaveta: "", prateleira: "", caixa: "",
  responsavel: "", dataCadastro: "", observacoes: "",
  fotos: [],
  tipoAmostra: "rocha"
});

// ---------- Classificação QAPF (Streckeisen 1976 plutônicas / 1978 vulcânicas) ----------
const num = (v) => { const n = parseFloat(String(v ?? "").replace(",", ".")); return Number.isFinite(n) && n > 0 ? n : 0; };
function classificarQAPF(q0, a0, p0, f0, modo) {
  let q = num(q0), a = num(a0), p = num(p0), f = num(f0);
  if (q + a + p + f <= 0) return null;
  let aviso = "";
  if (q > 0 && f > 0) {
    aviso = "Quartzo e feldspatoide não ocorrem juntos na mesma rocha. Foi considerado apenas o de maior porcentagem.";
    if (q >= f) f = 0; else q = 0;
  }
  const lado = f > 0 ? "F" : "Q";
  const tot = lado === "F" ? a + p + f : q + a + p;
  const Q = (q / tot) * 100, A = (a / tot) * 100, P = (p / tot) * 100, F = (f / tot) * 100;
  const nivel = lado === "F" ? F : Q;
  const t = a + p > 0 ? (p / (a + p)) * 100 : 50; // P/(A+P)
  const faixa = (lim, nomes) => { for (let i = 0; i < lim.length; i++) if (t < lim[i]) return nomes[i]; return nomes[nomes.length - 1]; };
  let nome = "";
  const vulc = modo === "Vulcânica";
  if (lado === "Q") {
    if (!vulc) {
      if (nivel > 90) nome = "Quartzolito";
      else if (nivel > 60) nome = "Granitoide rico em quartzo";
      else if (nivel > 20) nome = faixa([10, 35, 65, 90], ["Álcali feldspato granito", "Sienogranito (granito)", "Monzogranito (granito)", "Granodiorito", "Tonalito"]);
      else if (nivel > 5) nome = faixa([10, 35, 65, 90], ["Quartzo álcali feldspato sienito", "Quartzo sienito", "Quartzo monzonito", "Quartzo monzodiorito / quartzo monzogabro", "Quartzo diorito / quartzo gabro / quartzo anortosito"]);
      else nome = faixa([10, 35, 65, 90], ["Álcali feldspato sienito", "Sienito", "Monzonito", "Monzodiorito / monzogabro", "Diorito / gabro / anortosito"]);
    } else {
      if (nivel > 60) nome = "Rocha vulcânica rica em quartzo (Q > 60%)";
      else if (nivel > 20) nome = faixa([10, 65], ["Álcali feldspato riolito", "Riolito", "Dacito"]);
      else if (nivel > 5) nome = faixa([10, 35, 65], ["Álcali quartzo traquito", "Quartzo traquito", "Quartzo latito", "Andesito / Basalto"]);
      else nome = faixa([10, 35, 65], ["Álcali feldspato traquito", "Traquito", "Latito", "Andesito / Basalto"]);
    }
  } else {
    if (!vulc) {
      if (nivel > 60) nome = "Foidolito";
      else if (nivel > 10) nome = faixa([10, 50, 90], ["Feldspatoide sienito", "Feldspatoide monzosienito", "Feldspatoide monzodiorito / monzogabro", "Feldspatoide diorito / gabro"]);
      else nome = faixa([10, 35, 65, 90], ["Álcali feldspato sienito com feldspatoide", "Sienito com feldspatoide", "Monzonito com feldspatoide", "Monzodiorito / monzogabro com feldspatoide", "Diorito / gabro / anortosito com feldspatoide"]);
    } else {
      if (nivel > 90) nome = "Foidito";
      else if (nivel > 60) nome = t < 50 ? "Foidito fonolítico" : "Foidito tefrítico";
      else if (nivel > 10) nome = faixa([10, 50, 90], ["Fonolito", "Fonolito tefrítico", "Basanito fonolítico / Tefrito fonolítico", "Basanito (olivina > 10%) / Tefrito (olivina < 10%)"]);
      else nome = faixa([10, 35, 65], ["Álcali feldspato traquito com feldspatoide", "Traquito com feldspatoide", "Latito com feldspatoide", "Andesito / Basalto"]);
    }
  }
  return { nome, lado, Q, A, P, F, nivel, t, aviso };
}

// Campos do diagrama QAPF (mesmas faixas usadas em classificarQAPF). nivel = Q (parte de cima) ou F (parte de baixo); t = P/(A+P)
const QAPF_CAMPOS = (() => {
  const G = (lado, n1, n2, ts, nomes) => ts.slice(0, -1).map((t1, i) => ({ lado, n1, n2, t1, t2: ts[i + 1], nome: nomes[i] }));
  const T5 = [0, 10, 35, 65, 90, 100];
  const plut = [
    ...G("Q", 90, 100, [0, 100], ["Quartzolito"]),
    ...G("Q", 60, 90, [0, 100], ["Granitoide rico em quartzo"]),
    ...G("Q", 20, 60, T5, ["Álcali feldspato granito", "Sienogranito (granito)", "Monzogranito (granito)", "Granodiorito", "Tonalito"]),
    ...G("Q", 5, 20, T5, ["Quartzo álcali feldspato sienito", "Quartzo sienito", "Quartzo monzonito", "Quartzo monzodiorito / quartzo monzogabro", "Quartzo diorito / quartzo gabro / quartzo anortosito"]),
    ...G("Q", 0, 5, T5, ["Álcali feldspato sienito", "Sienito", "Monzonito", "Monzodiorito / monzogabro", "Diorito / gabro / anortosito"]),
    ...G("F", 0, 10, T5, ["Álcali feldspato sienito com feldspatoide", "Sienito com feldspatoide", "Monzonito com feldspatoide", "Monzodiorito / monzogabro com feldspatoide", "Diorito / gabro / anortosito com feldspatoide"]),
    ...G("F", 10, 60, [0, 10, 50, 90, 100], ["Feldspatoide sienito", "Feldspatoide monzosienito", "Feldspatoide monzodiorito / monzogabro", "Feldspatoide diorito / gabro"]),
    ...G("F", 60, 100, [0, 100], ["Foidolito"]),
  ];
  const vulc = [
    ...G("Q", 60, 100, [0, 100], ["Rocha vulcânica rica em quartzo (Q > 60%)"]),
    ...G("Q", 20, 60, [0, 10, 65, 100], ["Álcali feldspato riolito", "Riolito", "Dacito"]),
    ...G("Q", 5, 20, [0, 10, 35, 65, 100], ["Álcali quartzo traquito", "Quartzo traquito", "Quartzo latito", "Andesito / Basalto"]),
    ...G("Q", 0, 5, [0, 10, 35, 65, 100], ["Álcali feldspato traquito", "Traquito", "Latito", "Andesito / Basalto"]),
    ...G("F", 0, 10, [0, 10, 35, 65, 100], ["Álcali feldspato traquito com feldspatoide", "Traquito com feldspatoide", "Latito com feldspatoide", "Andesito / Basalto"]),
    ...G("F", 10, 60, [0, 10, 50, 90, 100], ["Fonolito", "Fonolito tefrítico", "Basanito fonolítico / Tefrito fonolítico", "Basanito (olivina > 10%) / Tefrito (olivina < 10%)"]),
    ...G("F", 60, 90, [0, 50, 100], ["Foidito fonolítico", "Foidito tefrítico"]),
    ...G("F", 90, 100, [0, 100], ["Foidito"]),
  ];
  const num = (l) => l.map((c, i) => ({ ...c, id: i + 1 }));
  return { "Plutônica": num(plut), "Vulcânica": num(vulc) };
})();
function findCampoQAPF(modo, lado, nivel, t) {
  return (QAPF_CAMPOS[modo] || []).find((c) => c.lado === lado && (c.n1 === 0 ? nivel >= 0 : nivel > c.n1) && nivel <= c.n2 && t >= c.t1 && (c.t2 === 100 ? t <= 100 : t < c.t2)) || null;
}
function nomeSugeridoQAPF(res, teorSilica) {
  if (!res) return "";
  if (res.nome === "Andesito / Basalto") {
    if ((teorSilica || "").startsWith("Básicas")) return "Basalto";
    if ((teorSilica || "").startsWith("Intermediárias")) return "Andesito";
  }
  return res.nome;
}

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
    <grupo>${escXML(m.grupo)}</grupo>
    <cor>${escXML(m.cor)}</cor>
    <traco>${escXML(m.traco)}</traco>
    <brilho>${escXML(m.brilho)}</brilho>
    <tipoBrilho>${escXML(m.brilhoDetalhe)}</tipoBrilho>
    <diafaneidade>${escXML(m.transparencia)}</diafaneidade>
    <clivagem>${escXML(m.clivagem)}</clivagem>
    <quantidadePlanos>${escXML(m.planosClivagem)}</quantidadePlanos>
    <reacaoAcido>${escXML(m.reacaoAcido)}</reacaoAcido>
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
    <estadoConservacao>${escXML(m.estadoConservacao)}</estadoConservacao>
    <registro>
        <dataCadastro>${escXML(m.dataCadastro)}</dataCadastro>
        <responsavel>${escXML(m.responsavel)}</responsavel>
        <observacoes>${escXML(m.observacoes)}</observacoes>
    </registro>
  </mineral>`;
}

function rochaEspecificoXML(r) {
  if (r.tipoRocha === "Ígnea") {
    return `    <especificoIgnea>
        <localizacao>
            <coordenadaX>${escXML(r.coordenadaX)}</coordenadaX>
            <coordenadaY>${escXML(r.coordenadaY)}</coordenadaY>
            <tipoAfloramento>${escXML(r.tipoAfloramento)}</tipoAfloramento>
        </localizacao>
        <corAlteracao>${escXML(r.corAlteracao)}</corAlteracao>
        <estruturasVulcanicas>
            <pequenaEscala>${escXML(r.estVulcPequena)}</pequenaEscala>
            <grandeEscala>${escXML(r.estVulcGrande)}</grandeEscala>
        </estruturasVulcanicas>
        <estruturasPlutonicas>
            <pequenaEscala>${escXML(r.estPlutPequena)}</pequenaEscala>
            <escalaMedia>${escXML(r.estPlutMedia)}</escalaMedia>
        </estruturasPlutonicas>
        <mineralogiaFelsica>
            <feldspatoPotassicoPct>${escXML(r.felFeldspatoK)}</feldspatoPotassicoPct>
            <plagioclasioPct>${escXML(r.felPlagioclasio)}</plagioclasioPct>
            <quartzoPct>${escXML(r.felQuartzo)}</quartzoPct>
            <muscovitaPct>${escXML(r.felMuscovita)}</muscovitaPct>
            <feldspatoidesPct>${escXML(r.felFeldspatoide)}</feldspatoidesPct>
            <outroMineral>${escXML(r.felOutroNome)}</outroMineral>
            <outroPct>${escXML(r.felOutroPct)}</outroPct>
        </mineralogiaFelsica>
        <mineralogiaMafica>
            <turmalinaPct>${escXML(r.mafTurmalina)}</turmalinaPct>
            <anfibolioPct>${escXML(r.mafAnfibolio)}</anfibolioPct>
            <biotitaPct>${escXML(r.mafBiotita)}</biotitaPct>
            <magnetitaPct>${escXML(r.mafMagnetita)}</magnetitaPct>
            <outroMineral>${escXML(r.mafOutroNome)}</outroMineral>
            <outroPct>${escXML(r.mafOutroPct)}</outroPct>
        </mineralogiaMafica>
        <mineraisAcessorios>${escXML(r.mineralogiaSecundaria)}</mineraisAcessorios>
        <textura>
            <grauCristalinidade>${escXML(r.grauCristalinidade)}</grauCristalinidade>
            <grauVisibilidade>${escXML(r.grauVisibilidade)}</grauVisibilidade>
            <tipoVisibilidade>${escXML(r.visibilidadeTipo)}</tipoVisibilidade>
            <granulacao>${escXML(r.tamanhoCristaisAbsoluto)}</granulacao>
            <formaGeometricaCristais>${escXML(r.formaGeometricaCristais)}</formaGeometricaCristais>
        </textura>
        <classificacao>
            <ambienteFormacao>${escXML(r.ambienteFormacaoIgnea)}</ambienteFormacao>
            <indiceCor>${escXML(r.indiceCorIgnea)}</indiceCor>
            <teorSilica>${escXML(r.saturacaoSilica)}</teorSilica>
            <nomenclatura>${escXML(r.nomenclaturaIgnea)}</nomenclatura>
        </classificacao>
    </especificoIgnea>\n`;
  }
  if (r.tipoRocha === "Sedimentar") {
    return `    <especificoSedimentar>
        <localizacao>
            <coordenadaX>${escXML(r.coordenadaX)}</coordenadaX>
            <coordenadaY>${escXML(r.coordenadaY)}</coordenadaY>
            <tipoAfloramento>${escXML(r.tipoAfloramento)}</tipoAfloramento>
        </localizacao>
        <aspectosGerais>
            <cor>${escXML(r.sedCor)}</cor>
            <origemCor>${escXML(r.sedCorOrigem)}</origemCor>
            <coesao>${escXML(r.sedCoesao)}</coesao>
            <porosidade>${escXML(r.sedPorosidade)}</porosidade>
            <reacaoHCl>${escXML(r.rochaReacaoHCl)}</reacaoHCl>
        </aspectosGerais>
        <textura>
            <materialFormacao>${escXML(r.sedMaterialFormacao)}</materialFormacao>
            <granulometria>${escXML(r.sedGranulometria)}</granulometria>
            <selecionamento>${escXML(r.sedSelecionamento)}</selecionamento>
            <arredondamento>${escXML(r.sedArredondamento)}</arredondamento>
            <esfericidade>${escXML(r.sedEsfericidade)}</esfericidade>
            <arcabouco>${escXML(r.sedArcabouco)}</arcabouco>
            <texturaNaoClastica>${escXML(r.sedTexturaNaoClastica)}</texturaNaoClastica>
            <cimento>${escXML(r.sedCimento)}</cimento>
        </textura>
        <estruturas>
            <acamamento>${escXML(r.sedAcamamento)}</acamamento>
            <espessuraCamadas>${escXML(r.sedEspessura)}</espessuraCamadas>
            <estruturasSuperficie>${escXML(r.sedEstrSuperficie)}</estruturasSuperficie>
            <fosseis>${escXML(r.sedFosseis)}</fosseis>
            <outrasPresente>${escXML(r.sedPossuiEstrutura)}</outrasPresente>
            <outrasDescricao>${escXML(r.sedEstruturaDescricao)}</outrasDescricao>
        </estruturas>
        <composicao>
            <quartzoPct>${escXML(r.sedQuartzo)}</quartzoPct>
            <feldspatoPct>${escXML(r.sedFeldspato)}</feldspatoPct>
            <liticosPct>${escXML(r.sedLiticos)}</liticosPct>
            <micasPct>${escXML(r.sedMicas)}</micasPct>
            <argilomineraisPct>${escXML(r.sedArgilominerais)}</argilomineraisPct>
            <calcitaPct>${escXML(r.sedCalcita)}</calcitaPct>
            <dolomitaPct>${escXML(r.sedDolomita)}</dolomitaPct>
            <oxidosFerroPct>${escXML(r.sedOxidosFe)}</oxidosFerroPct>
            <materiaOrganicaPct>${escXML(r.sedMateriaOrganica)}</materiaOrganicaPct>
            <outroMineral>${escXML(r.sedOutroNome)}</outroMineral>
            <outroPct>${escXML(r.sedOutroPct)}</outroPct>
            <mineraisAcessorios>${escXML(r.mineralogiaSecundaria)}</mineraisAcessorios>
        </composicao>
        <classificacao>
            <grupo>${escXML(r.sedClassificacao)}</grupo>
            <ambienteDeposicional>${escXML(r.sedAmbienteDep)}</ambienteDeposicional>
            <nomenclatura>${escXML(r.sedNomenclatura)}</nomenclatura>
        </classificacao>
    </especificoSedimentar>\n`;
  }
  if (r.tipoRocha === "Metamórfica") {
    return `    <especificoMetamorfica>
        <localizacao>
            <coordenadaX>${escXML(r.coordenadaX)}</coordenadaX>
            <coordenadaY>${escXML(r.coordenadaY)}</coordenadaY>
            <tipoAfloramento>${escXML(r.tipoAfloramento)}</tipoAfloramento>
        </localizacao>
        <aspectosGerais>
            <cor>${escXML(r.metCor)}</cor>
            <brilho>${escXML(r.metBrilho)}</brilho>
            <reacaoHCl>${escXML(r.rochaReacaoHCl)}</reacaoHCl>
        </aspectosGerais>
        <estrutura>
            <aspecto>${escXML(r.metAspecto)}</aspecto>
            <estruturasAssociadas>${escXML(r.metEstrAssoc)}</estruturasAssociadas>
        </estrutura>
        <textura>
            <granulacao>${escXML(r.metGranulacao)}</granulacao>
            <textura>${escXML(r.metTextura)}</textura>
            <formaCristais>${escXML(r.metForma)}</formaCristais>
            <porfiroblastos>${escXML(r.metPorfiroblastos)}</porfiroblastos>
        </textura>
        <mineralogia>
            <quartzoPct>${escXML(r.metQuartzo)}</quartzoPct>
            <feldspatoPct>${escXML(r.metFeldspato)}</feldspatoPct>
            <carbonatosPct>${escXML(r.metCarbonato)}</carbonatosPct>
            <muscovitaPct>${escXML(r.metMuscovita)}</muscovitaPct>
            <biotitaPct>${escXML(r.metBiotita)}</biotitaPct>
            <cloritaPct>${escXML(r.metClorita)}</cloritaPct>
            <anfibolioPct>${escXML(r.metAnfibolio)}</anfibolioPct>
            <piroxenioPct>${escXML(r.metPiroxenio)}</piroxenioPct>
            <granadaPct>${escXML(r.metGranada)}</granadaPct>
            <talcoSerpentinaPct>${escXML(r.metTalcoSerp)}</talcoSerpentinaPct>
            <grafitaPct>${escXML(r.metGrafita)}</grafitaPct>
            <aluminossilicatosPct>${escXML(r.metAluminossilicato)}</aluminossilicatosPct>
            <outroMineral>${escXML(r.metOutroNome)}</outroMineral>
            <outroPct>${escXML(r.metOutroPct)}</outroPct>
            <mineraisIndice>${escXML(r.mineraisIndice)}</mineraisIndice>
            <mineraisAcessorios>${escXML(r.mineralogiaSecundaria)}</mineraisAcessorios>
        </mineralogia>
        <classificacao>
            <tipoMetamorfismo>${escXML(r.tipoMetamorfismo)}</tipoMetamorfismo>
            <grauMetamorfico>${escXML(r.grauMetamorfico)}</grauMetamorfico>
            <facies>${escXML(r.metFacies)}</facies>
            <protolito>${escXML(r.rochaOriginal)}</protolito>
            <nomenclatura>${escXML(r.metNomenclatura)}</nomenclatura>
        </classificacao>
    </especificoMetamorfica>\n`;
  }
  return "";
}

function rochaToXML(r) {
  return `  <rocha>
    <codigo>${escXML(r.codigo)}</codigo>
    <nome>${escXML(r.nome)}</nome>
    <tipo>${escXML(r.tipoRocha)}</tipo>
${rochaEspecificoXML(r)}    <registro>
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
  const [tema, setTema] = useState(() => {
    try {
      const salvo = localStorage.getItem("geoacervo-tema");
      if (salvo === "dark" || salvo === "light") return salvo;
    } catch (e) { /* sem armazenamento */ }
    return window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
  });
  useEffect(() => {
    document.documentElement.setAttribute("data-theme", tema);
    try { localStorage.setItem("geoacervo-tema", tema); } catch (e) { /* ignora */ }
    const meta = document.querySelector('meta[name="theme-color"]');
    if (meta) meta.setAttribute("content", tema === "dark" ? "#0F1813" : "#0B3D2E");
  }, [tema]);
  const [mobileNavOpen, setMobileNavOpen] = useState(false);

  const [globalSearch, setGlobalSearch] = useState("");
  const [showGlobalResults, setShowGlobalResults] = useState(false);

  const [mineralForm, setMineralForm] = useState(null); // null = closed
  const [rochaForm, setRochaForm] = useState(null);
  const [viewItem, setViewItem] = useState(null); // ficha view/print
  const [toast, setToast] = useState(null);
  const [refreshing, setRefreshing] = useState(false);

  const printRef = useRef(null);
  const loadDadosRef = useRef(null);

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
    await loadDadosRef.current();
    return true;
  }, []);

  // ---------- Dados do acervo (minerais / rochas / histórico — cada registro é uma linha própria) ----------
  const loadDados = useCallback(async () => {
    const [{ data: m }, { data: r }, { data: h }] = await Promise.all([
      supabase.from("minerais").select("id, dados").order("created_at", { ascending: false }),
      supabase.from("rochas").select("id, dados").order("created_at", { ascending: false }),
      supabase.from("historico").select("*").order("created_at", { ascending: false }).limit(500),
    ]);
    setMinerais((m || []).map((row) => ({ ...row.dados, id: row.id })));
    setRochas((r || []).map((row) => ({ ...row.dados, id: row.id })));
    setHistorico((h || []).map((row) => ({ id: row.id, data: row.created_at, usuario: row.usuario, acao: row.acao, tipo: row.tipo, codigo: row.codigo, detalhes: row.detalhes })));
  }, []);
  loadDadosRef.current = loadDados;

  const loadAll = useCallback(async () => {
    try {
      const { data: { session } } = await supabase.auth.getSession();
      if (session?.user) await carregarPerfil(session.user); // já chama loadDados() internamente
    } catch (e) {
      setStorageError(true);
    }
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

  // Sincronização em tempo real: quando qualquer pessoa cadastra, edita ou
  // exclui um mineral/rocha, todo mundo com o site aberto recebe a
  // atualização na hora, sem precisar recarregar a página.
  useEffect(() => {
    if (!auth) return;

    const channel = supabase
      .channel("acervo-realtime")
      .on("postgres_changes", { event: "*", schema: "public", table: "minerais" }, (payload) => {
        if (payload.eventType === "INSERT") {
          const novo = { ...payload.new.dados, id: payload.new.id };
          setMinerais((prev) => (prev.some((m) => m.id === novo.id) ? prev : [novo, ...prev]));
        } else if (payload.eventType === "UPDATE") {
          const atualizado = { ...payload.new.dados, id: payload.new.id };
          setMinerais((prev) => prev.map((m) => (m.id === atualizado.id ? atualizado : m)));
        } else if (payload.eventType === "DELETE") {
          setMinerais((prev) => prev.filter((m) => m.id !== payload.old.id));
        }
      })
      .on("postgres_changes", { event: "*", schema: "public", table: "rochas" }, (payload) => {
        if (payload.eventType === "INSERT") {
          const novo = { ...payload.new.dados, id: payload.new.id };
          setRochas((prev) => (prev.some((r) => r.id === novo.id) ? prev : [novo, ...prev]));
        } else if (payload.eventType === "UPDATE") {
          const atualizado = { ...payload.new.dados, id: payload.new.id };
          setRochas((prev) => prev.map((r) => (r.id === atualizado.id ? atualizado : r)));
        } else if (payload.eventType === "DELETE") {
          setRochas((prev) => prev.filter((r) => r.id !== payload.old.id));
        }
      })
      .subscribe();

    return () => { supabase.removeChannel(channel); };
  }, [auth]);

  const showToast = (msg, type = "success") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3200);
  };

  const logHistory = useCallback(async (acao, tipo, codigo, detalhes) => {
    const { data, error } = await supabase.from("historico").insert({ usuario: auth?.nome || "—", acao, tipo, codigo, detalhes }).select().single();
    if (!error && data) {
      setHistorico((prev) => [{ id: data.id, data: data.created_at, usuario: data.usuario, acao: data.acao, tipo: data.tipo, codigo: data.codigo, detalhes: data.detalhes }, ...prev].slice(0, 500));
    }
  }, [auth]);

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

  const logout = async () => { await supabase.auth.signOut(); setAuth(null); setLoginForm({ u: "", p: "" }); setPage("dashboard"); setMinerais([]); setRochas([]); setHistorico([]); setUsuarios([]); };

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
    const isNew = !data.id || isDuplicate;
    if (isNew) {
      const codigo = isDuplicate ? proximoCodigo(minerais, "MIN") : (data.codigo || proximoCodigo(minerais, "MIN"));
      const payload = { ...data, codigo, dataCadastro: data.dataCadastro || todayISO(), responsavel: data.responsavel || auth.nome };
      delete payload.id;
      const { data: row, error } = await supabase.from("minerais").insert({ codigo, nome: payload.nome, dados: payload }).select("id, dados").single();
      if (error) return showToast("Erro ao salvar mineral: " + error.message, "error");
      const novo = { ...row.dados, id: row.id };
      setMinerais((prev) => [novo, ...prev]);
      logHistory(isDuplicate ? "Duplicado" : "Cadastrado", "Mineral", novo.codigo, novo.nome);
    } else {
      const { error } = await supabase.from("minerais").update({ codigo: data.codigo, nome: data.nome, dados: data, updated_at: new Date().toISOString() }).eq("id", data.id);
      if (error) return showToast("Erro ao salvar mineral: " + error.message, "error");
      setMinerais((prev) => prev.map((m) => (m.id === data.id ? data : m)));
      logHistory("Editado", "Mineral", data.codigo, data.nome);
    }
    setMineralForm(null);
    showToast(isNew ? "Mineral cadastrado com sucesso." : "Mineral atualizado com sucesso.");
  };

  const deleteMineral = async (item) => {
    if (!isAdmin) return showToast("Apenas administradores podem excluir registros.", "error");
    if (!confirm(`Excluir o mineral "${item.nome}" (${item.codigo})? Esta ação não pode ser desfeita.`)) return;
    const { error } = await supabase.from("minerais").delete().eq("id", item.id);
    if (error) return showToast("Erro ao excluir: " + error.message, "error");
    setMinerais((prev) => prev.filter((m) => m.id !== item.id));
    logHistory("Excluído", "Mineral", item.codigo, item.nome);
    showToast("Mineral excluído.");
  };

  // ---------- CRUD Rochas ----------
  const saveRocha = async (data, isDuplicate = false) => {
    if (!isAdmin) return showToast("Apenas administradores podem cadastrar ou editar rochas.", "error");
    const isNew = !data.id || isDuplicate;
    if (isNew) {
      const codigo = isDuplicate ? proximoCodigo(rochas, "ROC") : (data.codigo || proximoCodigo(rochas, "ROC"));
      const payload = { ...data, codigo, dataCadastro: data.dataCadastro || todayISO(), responsavel: data.responsavel || auth.nome };
      delete payload.id;
      const { data: row, error } = await supabase.from("rochas").insert({ codigo, nome: payload.nome, dados: payload }).select("id, dados").single();
      if (error) return showToast("Erro ao salvar rocha: " + error.message, "error");
      const novo = { ...row.dados, id: row.id };
      setRochas((prev) => [novo, ...prev]);
      logHistory(isDuplicate ? "Duplicado" : "Cadastrado", "Rocha", novo.codigo, novo.nome);
    } else {
      const { error } = await supabase.from("rochas").update({ codigo: data.codigo, nome: data.nome, dados: data, updated_at: new Date().toISOString() }).eq("id", data.id);
      if (error) return showToast("Erro ao salvar rocha: " + error.message, "error");
      setRochas((prev) => prev.map((r) => (r.id === data.id ? data : r)));
      logHistory("Editado", "Rocha", data.codigo, data.nome);
    }
    setRochaForm(null);
    showToast(isNew ? "Rocha cadastrada com sucesso." : "Rocha atualizada com sucesso.");
  };

  const deleteRocha = async (item) => {
    if (!isAdmin) return showToast("Apenas administradores podem excluir registros.", "error");
    if (!confirm(`Excluir a rocha "${item.nome}" (${item.codigo})? Esta ação não pode ser desfeita.`)) return;
    const { error } = await supabase.from("rochas").delete().eq("id", item.id);
    if (error) return showToast("Erro ao excluir: " + error.message, "error");
    setRochas((prev) => prev.filter((r) => r.id !== item.id));
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
      if (!confirm("Restaurar este backup vai ADICIONAR os minerais e rochas do arquivo ao acervo atual (não apaga o que já existe). Deseja continuar?")) return;
      if (Array.isArray(data.minerais) && data.minerais.length) {
        const linhas = data.minerais.map((m) => ({ codigo: m.codigo, nome: m.nome, dados: m }));
        const { error } = await supabase.from("minerais").insert(linhas);
        if (error) throw error;
      }
      if (Array.isArray(data.rochas) && data.rochas.length) {
        const linhas = data.rochas.map((r) => ({ codigo: r.codigo, nome: r.nome, dados: r }));
        const { error } = await supabase.from("rochas").insert(linhas);
        if (error) throw error;
      }
      await loadDadosRef.current();
      logHistory("Restauração de backup", "Sistema", "-", `Backup de ${data.exportadoEm || "data desconhecida"}`);
      showToast("Backup restaurado com sucesso.");
    } catch (e) { showToast("Erro ao restaurar backup: " + (e.message || "arquivo inválido"), "error"); }
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
      const classe = i.tipoAmostra === "mineral" ? i.grupo : i.tipoRocha;
      return [i.nome, i.codigo, i.cor, classe, i.tipoAmostra, local, municipio]
        .some((v) => (v || "").toLowerCase().includes(q));
    }).slice(0, 12);
  }, [globalSearch, allItems]);

  // ============================= RENDER =============================
  if (booting) {
    return (
      <div style={S.bootWrap}>
        <style>{GLOBAL_CSS}

</style>
        <div style={S.bootCard}>
          <div className="crystalSpin" style={S.crystal} />
          <p style={{ fontFamily: "Inter", color: "var(--txt-0b3d2e)", marginTop: 16, fontWeight: 600 }}>Carregando acervo geológico…</p>
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
              <div style={{ fontFamily: "Fraunces", fontWeight: 600, fontSize: 23, color: "var(--txt-0b3d2e)", lineHeight: 1.1 }}>Nova senha</div>
              <div style={{ fontFamily: "Inter", fontSize: 14.5, color: "var(--txt-5b7768)", letterSpacing: 0.3 }}>GeoAcervo — redefinição de senha</div>
            </div>
          </div>
          {newPassOk ? (
            <div style={S.regOkBox}>
              <CheckCircle size={26} color="var(--txt-2e7d5b)" />
              <p style={{ fontWeight: 700, color: "var(--txt-0b3d2e)", margin: "10px 0 4px" }}>Senha redefinida!</p>
              <p style={{ fontSize: 14.5, color: "var(--txt-5b7768)", lineHeight: 1.6 }}>Você já pode entrar com sua nova senha.</p>
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
            <LogoGeo size={52} />
            <div>
              <div style={{ fontFamily: "Fraunces", fontWeight: 600, fontSize: 23, color: "var(--txt-0b3d2e)", lineHeight: 1.1 }}>GeoAcervo</div>
              <div style={{ fontFamily: "Inter", fontSize: 14.5, color: "var(--txt-5b7768)", letterSpacing: 0.3 }}>Gestão de Minerais & Rochas</div>
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
                <CheckCircle size={26} color="var(--txt-2e7d5b)" />
                <p style={{ fontWeight: 700, color: "var(--txt-0b3d2e)", margin: "10px 0 4px" }}>E-mail enviado!</p>
                <p style={{ fontSize: 14.5, color: "var(--txt-5b7768)", lineHeight: 1.6 }}>Se esse e-mail estiver cadastrado, você vai receber um link para redefinir sua senha. Confira também a caixa de spam.</p>
                <button type="button" style={S.btnGhost} onClick={() => { setAuthView("login"); setForgotOk(false); }}>Voltar para o login</button>
              </div>
            ) : (
              <div style={{ marginTop: 20 }}>
                <p style={{ fontSize: 14.5, color: "var(--txt-5b7768)", lineHeight: 1.6, marginBottom: 14 }}>Informe o e-mail da sua conta. Enviaremos um link para você definir uma nova senha.</p>
                <label style={S.label}>E-mail</label>
                <input style={S.input} type="email" value={forgotEmail} onChange={(e) => setForgotEmail(e.target.value)} placeholder="seu@email.com" autoFocus onKeyDown={(e) => e.key === "Enter" && doForgotPassword()} />
                {forgotErr && <div style={S.loginErr}><AlertCircle size={15} /> {forgotErr}</div>}
                <button type="button" onClick={doForgotPassword} style={S.btnPrimary}>Enviar link de recuperação <ChevronRight size={16} /></button>
                <button type="button" style={S.forgotLink} onClick={() => setAuthView("login")}>Voltar para o login</button>
              </div>
            )
          ) : regOk ? (
            <div style={S.regOkBox}>
              <CheckCircle size={26} color="var(--txt-2e7d5b)" />
              <p style={{ fontWeight: 700, color: "var(--txt-0b3d2e)", margin: "10px 0 4px" }}>Cadastro enviado!</p>
              <p style={{ fontSize: 14.5, color: "var(--txt-5b7768)", lineHeight: 1.6 }}>Sua conta está aguardando aprovação do administrador. Você poderá entrar assim que ela for aprovada (e, se a confirmação de e-mail estiver ativa no projeto, após confirmar seu e-mail).</p>
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
              <p style={{ fontSize: 13.5, color: "var(--txt-a9bdb2)", marginTop: 12, lineHeight: 1.6 }}>Sua conta ficará com status "Pendente" até um administrador aprovar o acesso.</p>
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
          <LogoGeo size={36} />
          {sidebarOpen && <div className="geoBrandText"><strong>GeoAcervo</strong><small>Minerais · Rochas · Conhecimento</small></div>}
        </div>
        <nav style={S.nav}>
          <NavItem icon={LayoutDashboard} label="Visão geral" active={page === "dashboard"} open={sidebarOpen} onClick={() => { setPage("dashboard"); setMobileNavOpen(false); }} />
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
          <button style={{ ...S.logoutBtn, color: "#cfe0d6" }} onClick={() => setTema((t) => (t === "dark" ? "light" : "dark"))}>{tema === "dark" ? <Sun size={16} /> : <Moon size={16} />} {sidebarOpen && (tema === "dark" ? "Tema claro" : "Tema escuro")}</button>
          <button style={S.logoutBtn} onClick={logout}><LogOut size={16} /> {sidebarOpen && "Sair"}</button>
        </div>
      </aside>
      {mobileNavOpen && <div style={S.overlay} className="mobileOverlay" onClick={() => setMobileNavOpen(false)} />}

      {/* Main */}
      <div style={S.main} className="main">
        <header style={S.topbar} className="topbar">
          <div className="mobileBrand"><LogoGeo size={34} /><div><strong>GeoAcervo</strong><small>Minerais e Rochas</small></div></div>
          <button className="onlyMobile" style={S.menuBtn} onClick={() => setMobileNavOpen(true)} aria-label="Abrir menu"><Menu size={20} /></button>
          <div style={S.searchBox}>
            <Search size={16} color="var(--txt-5b7768)" />
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
                    {r.tipoAmostra === "mineral" ? <Gem size={14} color="var(--txt-2e7d5b)" /> : <Mountain size={14} color="#C9A227" />}
                    <span style={S.mono}>{r.codigo}</span>
                    <span style={{ fontWeight: 600 }}>{r.nome}</span>
                    <span style={{ color: "var(--txt-7c9186)", fontSize: 14.5 }}>{r.cor}</span>
                  </div>
                ))}
              </div>
            )}
          </div>
          <button
            style={S.refreshBtn}
            title="Atualizar dados agora"
            onClick={async () => { setRefreshing(true); await loadDadosRef.current?.(); setRefreshing(false); showToast("Dados atualizados."); }}
          >
            <RefreshCw size={15} className={refreshing ? "spinIcon" : ""} />
          </button>
          <div style={S.userChip}>
            <div style={S.userAvatar}>{auth.nome.slice(0, 1).toUpperCase()}</div>
            <div style={{ lineHeight: 1.2 }}>
              <div style={{ fontWeight: 600, fontSize: 15.5 }}>{auth.nome}</div>
              <div style={{ fontSize: 13.5, color: "var(--txt-5b7768)" }}>{auth.role}</div>
            </div>
          </div>
        </header>

        <main style={S.content}>
          {page === "dashboard" && (
            <Dashboard stats={stats} ultimos={ultimosCadastros} onNovoMineral={() => setMineralForm(emptyMineral())} onNovaRocha={() => setRochaForm(emptyRocha())} onControle={() => setPage("controle")} onRelatorios={() => setPage("relatorios")} onView={setViewItem} minerais={minerais} rochas={rochas} isAdmin={isAdmin} />
          )}
          {page === "minerais" && (
            <ListaPage
              title="Minerais" icon={Gem} color="var(--txt-2e7d5b)" items={minerais}
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

      <nav className="mobileBottomNav" aria-label="Navegação principal">
        <button className={page === "dashboard" ? "active" : ""} onClick={() => setPage("dashboard")}><LayoutDashboard size={20} /><span>Início</span></button>
        <button className={page === "minerais" ? "active" : ""} onClick={() => setPage("minerais")}><Gem size={20} /><span>Minerais</span></button>
        <button className={page === "rochas" ? "active" : ""} onClick={() => setPage("rochas")}><Mountain size={20} /><span>Rochas</span></button>
        <button className={mobileNavOpen ? "active" : ""} onClick={() => setMobileNavOpen(true)}><Menu size={20} /><span>Menu</span></button>
      </nav>

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
  const recentes = ultimos.slice(0, 5);
  const galeria = [...minerais.map((x) => ({ ...x, tipoAmostra: "mineral" })), ...rochas.map((x) => ({ ...x, tipoAmostra: "rocha" }))]
    .filter((x) => x.fotoPrincipal || (Array.isArray(x.fotos) && x.fotos.length) || (Array.isArray(x.fotosAdicionais) && x.fotosAdicionais.length))
    .slice(0, 4);
  const imagemDe = (item) => item.fotoPrincipal || item.fotos?.[0] || item.fotosAdicionais?.[0] || "";

  return (
    <div className="geoDashboard">
      <section className="geoWelcome">
        <div className="geoWelcomeText">
          <span className="geoEyebrow">SEU ACERVO DIGITAL</span>
          <h1>Olá! Bem-vindo ao GeoAcervo</h1>
          <p>Gerencie e consulte seu acervo geológico em um só lugar.</p>
          <span className="geoWelcomeLine" />
        </div>
        <div className="geoWelcomeArt" aria-hidden="true"><Mountain size={170} strokeWidth={0.8} /></div>
      </section>

      <div className="geoStatsGrid">
        <div className="geoStatCard"><div className="geoStatIcon mineral"><Gem size={24} /></div><div><span>Minerais cadastrados</span><strong>{stats.totalMinerais}</strong></div><Gem className="geoStatWatermark" size={34} /></div>
        <div className="geoStatCard"><div className="geoStatIcon rock"><Mountain size={24} /></div><div><span>Rochas cadastradas</span><strong>{stats.totalRochas}</strong></div><Mountain className="geoStatWatermark" size={34} /></div>
        <div className="geoStatCard"><div className="geoStatIcon recent"><Calendar size={24} /></div><div><span>Cadastros neste mês</span><strong>{stats.doMes}</strong></div><ArrowUpRight className="geoStatWatermark" size={34} /></div>
        <div className="geoStatCard"><div className="geoStatIcon total"><Boxes size={24} /></div><div><span>Total no acervo</span><strong>{stats.total}</strong></div><Boxes className="geoStatWatermark" size={34} /></div>
      </div>

      <div className="geoDashboardColumns">
        <div className="geoDashboardLeft">
          <section className="geoPanel geoQuickPanel">
            <div className="geoPanelTitle"><ZapIcon /> <h2>Acesso rápido</h2></div>
            <div className="geoQuickActions">
              {isAdmin && <button className="geoQuickButton dark" onClick={onNovoMineral}><Gem size={25} /><span>Cadastrar<br />Mineral</span><ArrowUpRight size={18} /></button>}
              {isAdmin && <button className="geoQuickButton green" onClick={onNovaRocha}><Mountain size={25} /><span>Cadastrar<br />Rocha</span><ArrowUpRight size={18} /></button>}
              {isAdmin && <button className="geoQuickButton soft" onClick={onControle}><ClipboardList size={22} /><span>Controle de<br />cadastros</span><ArrowUpRight size={18} /></button>}
              {isAdmin && <button className="geoQuickButton soft" onClick={onRelatorios}><BarChart3 size={22} /><span>Ver<br />relatórios</span><ArrowUpRight size={18} /></button>}
              {!isAdmin && <p className="geoQuickHint">Consulte os minerais e rochas pelas opções do menu lateral.</p>}
            </div>
          </section>

          <section className="geoPanel geoGalleryPanel">
            <div className="geoPanelTitle"><ImageIcon size={21} /><h2>Explore nosso acervo</h2><span className="geoPanelNote">{galeria.length ? `${galeria.length} amostras com foto` : "Fotos cadastradas"}</span></div>
            {galeria.length ? <div className="geoGalleryGrid">{galeria.map((item) => <button key={`${item.tipoAmostra}-${item.id || item.codigo}`} className="geoGalleryItem" onClick={() => onView(item)} title={`Ver ${item.nome}`}><img src={imagemDe(item)} alt={item.nome || "Amostra geológica"} /><span>{item.nome || "Amostra sem nome"}</span></button>)}</div> : <div className="geoGalleryEmpty"><ImageIcon size={25} /><span>As fotos das amostras cadastradas aparecerão aqui.</span></div>}
          </section>
        </div>

        <div className="geoDashboardRight">
          <section className="geoPanel geoRecentPanel">
            <div className="geoPanelTitle"><History size={21} /><h2>Adicionados recentemente</h2>{isAdmin && <button className="geoSeeAll" onClick={onControle}>Ver todos <ArrowUpRight size={15} /></button>}</div>
            {recentes.length === 0 ? <EmptyState text="Ainda não há amostras cadastradas. Use os atalhos para começar." /> : <div className="geoRecentList">
              {recentes.map((it) => {
                const foto = imagemDe(it);
                return <button className="geoRecentRow" key={it.id} onClick={() => onView(it)}>
                  <span className="geoRecentThumb">{foto ? <img src={foto} alt="" /> : it.tipoAmostra === "mineral" ? <Gem size={20} /> : <Mountain size={20} />}</span>
                  <span className="geoRecentName">{it.nome || "Sem nome"}<small>{it.codigo || "Sem código"}</small></span>
                  <span className="geoRecentType">{it.tipoAmostra === "mineral" ? "Mineral" : "Rocha"}</span>
                  <span className="geoRecentDate">{fmtDate(it.dataCadastro)}</span>
                  <ChevronRight size={17} className="geoRecentArrow" />
                </button>;
              })}
            </div>}
          </section>

          <section className="geoKnowledgeBanner">
            <div className="geoKnowledgeIcon"><Shield size={28} /></div>
            <div><h2>Conhecimento ao seu alcance</h2><p>Uma plataforma para armazenar, consultar e compartilhar informações sobre minerais e rochas.</p></div>
            <Mountain className="geoKnowledgeArt" size={116} strokeWidth={0.8} />
          </section>
        </div>
      </div>
    </div>
  );
}

function ZapIcon() { return <span className="geoZap">ϟ</span>; }

function LogoGeo({ size = 40 }) {
  return (
    <svg width={size} height={size} viewBox="0 0 100 100" role="img" aria-label="GeoAcervo" style={{ flexShrink: 0, display: "block" }}>
      <rect width="100" height="100" rx="22" fill="#0b3d2e" />
      <polygon points="50,14 72,34 64,78 36,78 28,34" fill="#2E7D5B" />
      <polygon points="50,14 72,34 50,44" fill="#5fc08f" />
      <polygon points="50,14 28,34 50,44" fill="#3f9d73" />
      <polygon points="50,44 72,34 64,78" fill="#1f6a4a" />
      <polygon points="50,44 28,34 36,78" fill="#2a8059" />
      <polygon points="50,44 64,78 36,78" fill="#175a3d" />
    </svg>
  );
}

function QuickAction({ icon: Icon, label, onClick, color }) {
  return (
    <button style={S.quickCard} onClick={onClick}>
      <div style={{ ...S.quickIcon, background: color }}><Icon size={18} color="#fff" /></div>
      <span>{label}</span>
      <ArrowUpRight size={15} color="var(--txt-7c9186)" style={{ marginLeft: "auto" }} />
    </button>
  );
}

function TypeBadge({ item }) {
  const isMin = item.tipoAmostra === "mineral";
  return (
    <span style={{ ...S.badge, background: isMin ? "var(--bg-e4f2e9)" : "var(--bg-fbf2da)", color: isMin ? "var(--txt-0b3d2e)" : "#7a5f0a" }}>
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
      <Icon size={30} color="var(--txt-a9bdb2)" />
      <p>{text}</p>
    </div>
  );
}

function RestritoPage() {
  return (
    <div style={S.card}>
      <div style={S.emptyState}>
        <Lock size={30} color="var(--txt-a9bdb2)" />
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
        isAdmin && <button style={S.btnPrimarySm} onClick={onNew}><Plus size={16} /> {title === "Rochas" ? "Nova Rocha" : "Novo Mineral"}</button>
      } />
      <div style={S.filterRow}>
        <div style={S.searchBoxSm}><Search size={15} color="var(--txt-5b7768)" /><input style={S.searchInputSm} placeholder={`Buscar por nome, código ou ${badgeField === "grupo" ? "grupo" : "tipo"}…`} value={q} onChange={(e) => setQ(e.target.value)} /></div>
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
        {foto ? <img src={foto} alt={item.nome} style={S.sampleImgTag} /> : <div style={S.sampleImgPh}><ImageIcon size={22} color="var(--txt-a9bdb2)" /></div>}
        <span style={{ ...S.codeChip, borderColor: color }}>{item.codigo}</span>
      </div>
      <div style={{ padding: "12px 14px 14px" }}>
        <div style={{ fontFamily: "Fraunces", fontWeight: 600, fontSize: 17.5, color: "var(--txt-0b3d2e)" }}>{item.nome || "(sem nome)"}</div>
        <div style={{ fontSize: 14.5, color: "var(--txt-5b7768)", marginTop: 2 }}>{isMin ? (item.grupo || "—") : (item.tipoRocha || "—")} · {item.cor || "sem cor definida"}</div>
        <div style={{ display: "flex", gap: 6, marginTop: 12, flexWrap: "wrap" }}>
          <button style={S.iconBtn} title="Visualizar" onClick={() => onView(item)}><Eye size={14} /></button>
          <button style={S.iconBtn} title="Imprimir ficha" onClick={() => onPrint(item)}><Printer size={14} /></button>
          <button style={S.iconBtn} title="Exportar XML" onClick={() => onExportXML(item)}><FileText size={14} /></button>
          {isAdmin && <>
            <button style={S.iconBtn} title="Editar" onClick={() => onEdit(item)}><Edit size={14} /></button>
            <button style={S.iconBtn} title="Duplicar" onClick={() => onDuplicate(item)}><Copy size={14} /></button>
            <button style={{ ...S.iconBtn, marginLeft: "auto", color: "var(--txt-8a2f2f)" }} title="Excluir" onClick={() => onDelete(item)}><Trash2 size={14} /></button>
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
          <div><label style={{ ...S.labelSm, minHeight: 38, display: "flex", alignItems: "flex-end" }}>Nome</label><input style={S.inputSm} value={nome} onChange={(e) => setNome(e.target.value)} placeholder="Buscar por nome…" /></div>
          <div><label style={{ ...S.labelSm, minHeight: 38, display: "flex", alignItems: "flex-end" }}>Código</label><input style={S.inputSm} value={codigo} onChange={(e) => setCodigo(e.target.value)} placeholder="Ex: MIN001" /></div>
          <div><label style={{ ...S.labelSm, minHeight: 38, display: "flex", alignItems: "flex-end" }}>Tipo</label>
            <select style={S.inputSm} value={tipo} onChange={(e) => { setTipo(e.target.value); setGrupo(""); }}>
              <option value="">Todos</option><option value="mineral">Mineral</option><option value="rocha">Rocha</option>
            </select>
          </div>
          <div><label style={{ ...S.labelSm, minHeight: 38, display: "flex", alignItems: "flex-end" }}>Grupo mineral / Tipo de rocha</label>
            <select style={S.inputSm} value={grupo} onChange={(e) => setGrupo(e.target.value)}>
              <option value="">Todos</option>
              {(tipo === "rocha" ? TIPOS_ROCHA : tipo === "mineral" ? GRUPOS_MINERALOGICOS : [...GRUPOS_MINERALOGICOS, ...TIPOS_ROCHA]).map((g) => <option key={g} value={g}>{g}</option>)}
            </select>
          </div>
          <div><label style={{ ...S.labelSm, minHeight: 38, display: "flex", alignItems: "flex-end" }}>Data inicial</label><input type="date" style={S.inputSm} value={dataIni} onChange={(e) => setDataIni(e.target.value)} /></div>
          <div><label style={{ ...S.labelSm, minHeight: 38, display: "flex", alignItems: "flex-end" }}>Data final</label><input type="date" style={S.inputSm} value={dataFim} onChange={(e) => setDataFim(e.target.value)} /></div>
          <div><label style={{ ...S.labelSm, minHeight: 38, display: "flex", alignItems: "flex-end" }}>Responsável</label>
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
                        <button style={{ ...S.iconBtn, color: "var(--txt-8a2f2f)" }} title="Excluir" onClick={() => onDelete(it)}><Trash2 size={13} /></button>
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
                    <td style={{ color: "var(--txt-5b7768)" }}>{h.detalhes}</td>
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
    const cfg = status === "Pendente" ? { bg: "var(--bg-fbf2da)", c: "var(--txt-7a5f0a)" } : status === "Recusado" ? { bg: "var(--bg-fbeaea)", c: "var(--txt-8a2f2f)" } : { bg: "var(--bg-e4f2e9)", c: "var(--txt-0b3d2e)" };
    return <span style={{ ...S.badge, background: cfg.bg, color: cfg.c }}>{status || "Aprovado"}</span>;
  };
  return (
    <div>
      <PageHeader title="Usuários e Permissões" subtitle="Contas criadas pelos próprios usuários na tela de login (aba 'Criar conta') e aprovadas aqui" />

      <div style={{ ...S.card, borderColor: "var(--bd-e8dcae)", background: "var(--bg-fffcf2)" }}>
        <p style={{ fontSize: 14.5, color: "var(--txt-6b5817)", margin: 0, lineHeight: 1.7 }}>
          As senhas agora são gerenciadas com segurança pelo Supabase Auth (criptografadas), então não é possível
          cadastrar usuários manualmente por aqui nem ver/alterar a senha de ninguém — cada pessoa cria a própria
          conta pela tela de login. Você pode aprovar, recusar, promover a Administrador ou revogar o acesso.
        </p>
      </div>

      {pendentes.length > 0 && (
        <div style={{ ...S.card, borderColor: "var(--bd-e8dcae)", background: "var(--bg-fffcf2)" }}>
          <h3 style={S.cardTitle}><TriangleAlert size={15} color="var(--txt-9a7b0a)" style={{ verticalAlign: -2 }} /> {pendentes.length} solicitação(ões) de cadastro aguardando aprovação</h3>
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
                        <button style={{ ...S.iconBtn, background: "var(--bg-e4f2e9)", color: "var(--txt-0b3d2e)" }} title="Aprovar" onClick={() => onSave({ ...u, status: "Aprovado" })}><CheckCircle size={14} /></button>
                        <button style={{ ...S.iconBtn, background: "var(--bg-fbeaea)", color: "var(--txt-8a2f2f)" }} title="Recusar" onClick={() => onSave({ ...u, status: "Recusado" })}><X size={14} /></button>
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
                  <td><span style={{ ...S.badge, background: u.role === "Administrador" ? "var(--bg-e4f2e9)" : "var(--bg-efefef)", color: u.role === "Administrador" ? "#0B3D2E" : "#555" }}>{u.role === "Administrador" && <Shield size={11} />} {u.role}</span></td>
                  <td>{statusBadge(u.status)}</td>
                  <td>
                    <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                      <button style={S.iconBtn} title="Editar perfil/status" onClick={() => setForm(u)}><Edit size={14} /></button>
                      <button style={{ ...S.iconBtn, color: "var(--txt-8a2f2f)" }} title="Revogar acesso" onClick={() => onRevoke(u)} disabled={u.id === authId}><Lock size={14} /></button>
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
          <p style={{ color: "var(--txt-5b7768)", fontSize: 15.5, lineHeight: 1.6 }}>Gera um arquivo com todos os {stats.total} registros, usuários e histórico de alterações. Guarde em local seguro.</p>
          <button style={S.btnPrimary} onClick={onBackup}><Database size={16} /> Gerar backup agora</button>
        </div>
        <div style={S.card}>
          <h3 style={S.cardTitle}><Upload size={16} style={{ verticalAlign: -3 }} /> Restaurar backup</h3>
          <p style={{ color: "var(--txt-5b7768)", fontSize: 15.5, lineHeight: 1.6 }}>{isAdmin ? "Adiciona os minerais e rochas do arquivo ao acervo atual (não apaga o que já existe)." : "Apenas administradores podem restaurar backups."}</p>
          <input ref={fileRef} type="file" accept=".json" style={{ display: "none" }} onChange={(e) => e.target.files[0] && onRestore(e.target.files[0])} />
          <button style={S.btnGhost} disabled={!isAdmin} onClick={() => fileRef.current?.click()}><FileUp size={16} /> Selecionar arquivo .json</button>
        </div>
      </div>
      <div style={{ ...S.card, borderColor: "var(--bd-e8dcae)", background: "var(--bg-fffcf2)" }}>
        <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
          <TriangleAlert size={18} color="var(--txt-9a7b0a)" style={{ flexShrink: 0, marginTop: 2 }} />
          <p style={{ color: "var(--txt-6b5817)", fontSize: 15, lineHeight: 1.7, margin: 0 }}>
            Este ambiente de demonstração armazena os dados em um banco de dados chave-valor compartilhado entre os usuários do sistema. Para produção em escala (10.000+ amostras, múltiplos laboratórios, autenticação segura), recomenda-se migrar para um banco relacional dedicado (PostgreSQL/MySQL) com backend próprio — a estrutura de dados aqui já foi projetada para essa migração direta.
          </p>
        </div>
      </div>
    </div>
  );
}


function WizardProgress({ steps, current }) {
  const pct = ((current + 1) / steps.length) * 100;
  return (
    <div className="geoWizardProgress">
      <div className="geoWizardTop"><span>Etapa {current + 1} de {steps.length}</span><strong>{steps[current]}</strong></div>
      <div className="geoWizardTrack"><span style={{ width: `${pct}%` }} /></div>
      <div className="geoWizardSteps">{steps.map((name, i) => (
        <span key={name} className={i < current ? "done" : i === current ? "active" : ""}>{i + 1}</span>
      ))}</div>
    </div>
  );
}

function WizardActions({ current, total, onBack, onNext, onCancel, onSave, saveLabel }) {
  return (
    <div className="geoWizardActions">
      <button type="button" style={S.btnGhost} onClick={current === 0 ? onCancel : onBack}>
        {current === 0 ? "Cancelar" : <><ChevronLeft size={16} /> Voltar</>}
      </button>
      {current < total - 1 ? (
        <button type="button" style={{ ...S.btnPrimary, width: "auto", marginTop: 0, padding: "11px 22px" }} onClick={onNext}>
          Continuar <ChevronRight size={16} />
        </button>
      ) : (
        <button type="button" style={{ ...S.btnPrimary, width: "auto", marginTop: 0, padding: "11px 22px" }} onClick={onSave}>
          <Save size={16} /> {saveLabel}
        </button>
      )}
    </div>
  );
}

// ---------- Formulários ----------
function Field({ label, children, span }) {
  return <div className="geoFormField" style={{ gridColumn: span ? `span ${span}` : undefined }}><label className="geoFormFieldLabel" style={S.labelSm}>{label}</label>{children}</div>;
}

function ModalShell({ title, onClose, children, wide }) {
  return (
    <div style={S.modalOverlay} onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className={wide ? "geoFormModal" : ""} style={{ ...S.modal, maxWidth: wide ? 960 : 620 }}>
        <div style={S.modalHead}>
          <h3 style={{ fontFamily: "Fraunces", fontSize: 20, color: "var(--txt-0b3d2e)", margin: 0 }}>{title}</h3>
          <button style={S.iconBtn} onClick={onClose}><X size={17} /></button>
        </div>
        <div className={wide ? "geoFormModalBody" : ""} style={S.modalBody}>{children}</div>
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
        <Upload size={16} color="var(--txt-5b7768)" /> <span>Clique para enviar {multiple ? "imagens" : "uma imagem"}</span>
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
  const n = parseFloat(String(value ?? "").replace(",", "."));
  const valido = !isNaN(n);
  const fmt = (x) => String(x).replace(".", ",");
  const ajustar = (delta) => {
    const base = valido ? n : 0;
    const novo = Math.min(10, Math.max(0.5, Math.round((base + delta) * 2) / 2));
    onChange(String(novo));
  };
  const ref = valido ? MOHS_REF.find((m) => m.d === n) : null;
  const abaixo = valido && !ref ? [...MOHS_REF].reverse().find((m) => m.d < n) : null;
  const acima = valido && !ref ? MOHS_REF.find((m) => m.d > n) : null;
  const legenda = ref ? ref.m : abaixo && acima ? `entre ${abaixo.m} e ${acima.m}` : "";
  return (
    <div>
      <label style={S.labelSm}>Dureza (Escala de Mohs) {valido && <span style={{ color: "var(--txt-0b3d2e)", fontWeight: 700 }}>— {fmt(n)}{legenda ? ` (${legenda})` : ""}</span>}</label>
      <div style={S.mohsRow}>
        {MOHS_REF.map((m) => (
          <button key={m.d} type="button" onClick={() => onChange(String(m.d))} title={`${m.d} — ${m.m}`}
            style={{ ...S.mohsFacet, background: valido && n >= m.d ? "#2E7D5B" : valido && n > m.d - 1 ? "#8fc4a8" : "var(--bg-e4f2e9)", color: valido && n >= m.d ? "var(--bg-fff)" : "#5b7768" }}>
            {m.d}
          </button>
        ))}
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 8 }}>
        <button type="button" aria-label="Diminuir 0,5" onClick={() => ajustar(-0.5)} style={{ ...S.input, width: 44, padding: 0, cursor: "pointer", fontWeight: 700 }}>−</button>
        <input type="text" inputMode="decimal" aria-label="Dureza de Mohs (aceita meios pontos, ex.: 5,5)" placeholder="Ex.: 5,5"
          value={value ?? ""} onChange={(e) => onChange(e.target.value.replace(/[^0-9.,]/g, "").slice(0, 4))}
          onBlur={() => { if (valido) onChange(String(Math.min(10, Math.max(0.5, Math.round(n * 2) / 2)))); }}
          style={{ ...S.input, width: 90, textAlign: "center" }} />
        <button type="button" aria-label="Aumentar 0,5" onClick={() => ajustar(0.5)} style={{ ...S.input, width: 44, padding: 0, cursor: "pointer", fontWeight: 700 }}>+</button>
        <span style={{ fontSize: 14, color: "#5b7768" }}>passos de 0,5 (0,5 a 10)</span>
      </div>
    </div>
  );
}

function MineralFormModal({ data, onClose, onSave, auth }) {
  const [f, setF] = useState({ ...data, dataCadastro: data.dataCadastro || todayISO(), responsavel: data.responsavel || auth.nome });
  const [step, setStep] = useState(0);
  const steps = ["Identificação", "Características", "Complementos", "Fotos e revisão"];
  const set = (k) => (e) => setF({ ...f, [k]: e?.target ? e.target.value : e });
  const submit = () => { if (!f.nome) { setStep(0); return; } onSave(f); };
  return (
    <ModalShell title={f.id ? `Editar mineral — ${f.codigo}` : "Novo Mineral"} onClose={onClose} wide>
      <div>
        <div className="geoMobileFormIntro"><span className="geoMobileFormIntroIcon"><Gem size={19} /></span><div><strong>Cadastro de mineral</strong><small>Preencha uma etapa por vez. Seus dados ficam guardados enquanto avança.</small></div></div>
        <WizardProgress steps={steps} current={step} />
        <div className="geoWizardPanel" style={{ display: step === 0 ? "block" : "none" }}>
        <SectionTitle icon={Gem}>Identificação</SectionTitle>
        <div style={S.formGrid3}>
          <Field label="Código"><input style={S.input} value={f.codigo} onChange={set("codigo")} placeholder="Automático se vazio" /></Field>
          <Field label="Nome do mineral *"><input style={S.input} value={f.nome} onChange={set("nome")} required /></Field>
          <Field label="Grupo mineral"><select style={S.input} value={f.grupo} onChange={set("grupo")}><option value="">Selecione…</option>{GRUPOS_MINERALOGICOS.map((g) => <option key={g}>{g}</option>)}</select></Field>
          <Field label="Composição química"><input style={S.input} value={f.composicaoQuimica} onChange={set("composicaoQuimica")} placeholder="Ex: SiO₂" /></Field>
        </div>

        </div>
        <div className="geoWizardPanel" style={{ display: step === 1 ? "block" : "none" }}>
        <SectionTitle icon={ScanLine}>Características físicas</SectionTitle>
        <div style={S.formGrid3}>
          <Field label="Cor"><input style={S.input} value={f.cor} onChange={set("cor")} /></Field>
          <Field label="Traço"><input style={S.input} value={f.traco} onChange={set("traco")} /></Field>
          <Field label="Brilho"><select style={S.input} value={f.brilho} onChange={set("brilho")}><option value="">Selecione…</option>{BRILHOS.map((g) => <option key={g}>{g}</option>)}</select></Field>
          {f.brilho === "Não metálico" && (
            <Field label="Tipo de brilho (não metálico)"><select style={S.input} value={f.brilhoDetalhe} onChange={set("brilhoDetalhe")}><option value="">Selecione…</option>{BRILHOS_NAO_METALICOS.map((g) => <option key={g}>{g}</option>)}</select></Field>
          )}
          <Field label="Diafaneidade"><select style={S.input} value={f.transparencia} onChange={set("transparencia")}><option value="">Selecione…</option>{TRANSPARENCIAS.map((g) => <option key={g}>{g}</option>)}</select></Field>
          <Field label="Clivagem"><select style={S.input} value={f.clivagem} onChange={set("clivagem")}><option value="">Selecione…</option>{CLIVAGENS.map((g) => <option key={g}>{g}</option>)}</select></Field>
          <Field label="Quantidade de planos (clivagem)"><input style={S.input} type="number" min="0" step="1" value={f.planosClivagem} onChange={set("planosClivagem")} /></Field>
          <Field label="Fratura"><select style={S.input} value={f.fratura} onChange={set("fratura")}><option value="">Selecione…</option>{FRATURAS.map((g) => <option key={g}>{g}</option>)}</select></Field>
          <Field label="Hábito cristalino"><input style={S.input} value={f.habito} onChange={set("habito")} /></Field>
          <Field label="Sistema cristalino"><select style={S.input} value={f.sistemaCristalino} onChange={set("sistemaCristalino")}><option value="">Selecione…</option>{SISTEMAS_CRISTALINOS.map((g) => <option key={g}>{g}</option>)}</select></Field>
          <Field label="Densidade (g/cm³)"><input style={S.input} type="number" step="0.01" value={f.densidade} onChange={set("densidade")} /></Field>
          <Field label="Magnetismo"><select style={S.input} value={f.magnetismo} onChange={set("magnetismo")}><option>Não</option><option>Sim</option><option>Fraco</option></select></Field>
          <Field label="Fluorescência"><select style={S.input} value={f.fluorescencia} onChange={set("fluorescencia")}><option>Não</option><option>Sim</option></select></Field>
          <Field label="Tenacidade"><select style={S.input} value={f.tenacidade} onChange={set("tenacidade")}><option value="">Selecione…</option>{TENACIDADES.map((g) => <option key={g}>{g}</option>)}</select></Field>
          <Field label="Reação com ácido"><select style={S.input} value={f.reacaoAcido} onChange={set("reacaoAcido")}><option value="">Selecione…</option>{REACAO_ACIDO.map((g) => <option key={g}>{g}</option>)}</select></Field>
        </div>
        <div style={{ marginTop: 12 }}><MohsSlider value={f.dureza} onChange={(v) => setF({ ...f, dureza: v })} /></div>

        </div>
        <div className="geoWizardPanel" style={{ display: step === 2 ? "block" : "none" }}>
        <SectionTitle icon={FileText}>Informações complementares</SectionTitle>
        <div style={S.formGrid2}>
          <Field label="Ambiente de formação"><input style={S.input} value={f.ambienteFormacao} onChange={set("ambienteFormacao")} /></Field>
          <Field label="Principais usos"><input style={S.input} value={f.usos} onChange={set("usos")} /></Field>
          <Field label="Estado de conservação"><select style={S.input} value={f.estadoConservacao} onChange={set("estadoConservacao")}><option value="">Selecione…</option>{ESTADOS_CONSERVACAO.map((g) => <option key={g}>{g}</option>)}</select></Field>
        </div>

        </div>
        <div className="geoWizardPanel" style={{ display: step === 2 ? "block" : "none" }}>
        <SectionTitle icon={ClipboardList}>Registro</SectionTitle>
        <div style={S.formGrid3}>
          <Field label="Data do cadastro"><input type="date" style={S.input} value={f.dataCadastro} onChange={set("dataCadastro")} /></Field>
          <Field label="Responsável pelo cadastro"><input style={S.input} value={f.responsavel} onChange={set("responsavel")} /></Field>
          <Field label="Observações" span={3}><textarea style={{ ...S.input, minHeight: 60, height: "auto", padding: "11px 13px" }} value={f.observacoes} onChange={set("observacoes")} /></Field>
        </div>

        </div>
        <div className="geoWizardPanel" style={{ display: step === 3 ? "block" : "none" }}>
        <SectionTitle icon={ImageIcon}>Mídia</SectionTitle>
        <div style={S.formGrid2}>
          <PhotoPicker label="Foto principal" value={f.fotoPrincipal} onChange={(v) => setF({ ...f, fotoPrincipal: v })} />
          <PhotoPicker label="Fotos adicionais (até 6)" value={f.fotosAdicionais} onChange={(v) => setF({ ...f, fotosAdicionais: v })} multiple />
        </div>

        </div>
        <WizardActions current={step} total={steps.length}
          onBack={() => setStep((v) => Math.max(0, v - 1))}
          onNext={() => setStep((v) => Math.min(steps.length - 1, v + 1))}
          onCancel={onClose} onSave={submit} saveLabel="Salvar mineral" />
      </div>
    </ModalShell>
  );
}

function SubTitle({ children }) {
  return <div style={{ marginTop: 14, fontSize: 13.5, fontWeight: 700, color: "var(--txt-5b7768)", textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 8 }}>{children}</div>;
}

function SomaPct({ f, keys }) {
  const soma = keys.reduce((t, k) => t + num(f[k]), 0);
  if (!soma) return null;
  const alto = soma > 100.0001;
  return <div style={{ marginTop: 6, fontSize: 14, fontWeight: 600, color: alto ? "#b3261e" : "var(--txt-5b7768)" }}>Soma: {String(Math.round(soma * 100) / 100).replace(".", ",")}%{alto ? " — passou de 100%, confira os valores" : ""}</div>;
}

function SugestaoNome({ sugestao, atual, onUse }) {
  if (!sugestao || sugestao === atual) return null;
  return (
    <div style={{ marginTop: 6, fontSize: 14, color: "var(--txt-5b7768)" }}>
      Sugestão pelos dados informados: <b>{sugestao}</b>{" "}
      <button type="button" onClick={() => onUse(sugestao)} style={{ border: "1px solid var(--bd-dfe8e2)", background: "transparent", color: "var(--txt-2e7d5b)", borderRadius: 6, padding: "2px 9px", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>Usar</button>
    </div>
  );
}

function RochaFormModal({ data, onClose, onSave, auth }) {
  const [f, setF] = useState({ ...data, dataCadastro: data.dataCadastro || todayISO(), responsavel: data.responsavel || auth.nome });
  const [step, setStep] = useState(0);
  const steps = ["Identificação", "Análise da rocha", "Registro e fotos"];
  const set = (k) => (e) => setF({ ...f, [k]: e?.target ? e.target.value : e });
  const submit = () => { if (!f.nome || !f.tipoRocha) { setStep(0); return; } onSave(f); };
  // Helpers de campo (funções que devolvem JSX, para não perder o foco ao digitar)
  const sel = (label, k, opts, extra) => {
    const lista = f[k] && !opts.includes(f[k]) ? [f[k], ...opts] : opts; // mantém valores de registros antigos
    return <Field label={label}><select style={S.input} value={f[k]} onChange={extra ? (e) => extra(e.target.value) : set(k)}><option value="">Selecione…</option>{lista.map((g) => <option key={g}>{g}</option>)}</select></Field>;
  };
  const pct = (label, k) => <Field label={label}><input style={S.input} type="number" min="0" max="100" value={f[k]} onChange={set(k)} placeholder="%" /></Field>;
  const txt = (label, k, ph) => <Field label={label}><input style={S.input} value={f[k]} onChange={set(k)} placeholder={ph} /></Field>;
  const clastica = !f.sedMaterialFormacao || f.sedMaterialFormacao.startsWith("Clástica");
  const SED_PCT = ["sedQuartzo", "sedFeldspato", "sedLiticos", "sedMicas", "sedArgilominerais", "sedCalcita", "sedDolomita", "sedOxidosFe", "sedMateriaOrganica", "sedOutroPct"];
  const MET_PCT = ["metQuartzo", "metFeldspato", "metCarbonato", "metMuscovita", "metBiotita", "metClorita", "metAnfibolio", "metPiroxenio", "metGranada", "metTalcoSerp", "metGrafita", "metAluminossilicato", "metOutroPct"];
  return (
    <ModalShell title={f.id ? `Editar rocha — ${f.codigo}` : "Nova Rocha"} onClose={onClose} wide>
      <div>
        <div className="geoMobileFormIntro"><span className="geoMobileFormIntroIcon"><Mountain size={19} /></span><div><strong>Cadastro de rocha</strong><small>As etapas se adaptam ao tipo de rocha selecionado.</small></div></div>
        <WizardProgress steps={steps} current={step} />
        <div className="geoWizardPanel" style={{ display: step === 0 ? "block" : "none" }}>
        <SectionTitle icon={Mountain}>Identificação</SectionTitle>
        <div style={S.formGrid3}>
          <Field label="Código"><input style={S.input} value={f.codigo} onChange={set("codigo")} placeholder="Automático se vazio" /></Field>
          <Field label="Nome *"><input style={S.input} value={f.nome} onChange={set("nome")} required /></Field>
          <Field label="Tipo"><select style={S.input} value={f.tipoRocha} onChange={set("tipoRocha")}><option value="">Selecione…</option>{TIPOS_ROCHA.map((g) => <option key={g}>{g}</option>)}</select></Field>
        </div>

        </div>
        <div className="geoWizardPanel" style={{ display: step === 1 ? "block" : "none" }}>
        {!f.tipoRocha && (
          <div style={S.tipoHint}><AlertCircle size={14} /> Selecione o "Tipo" na Identificação acima para ver os campos específicos de análise dessa rocha.</div>
        )}

        {f.tipoRocha === "Ígnea" && (
          <>
            <SectionTitle icon={Mountain}>Ígnea — 1. Localização</SectionTitle>
            <div style={S.formGrid3}>
              <Field label="Coordenada X"><input style={S.input} value={f.coordenadaX} onChange={set("coordenadaX")} /></Field>
              <Field label="Coordenada Y"><input style={S.input} value={f.coordenadaY} onChange={set("coordenadaY")} /></Field>
              <Field label="Tipo de afloramento"><input style={S.input} value={f.tipoAfloramento} onChange={set("tipoAfloramento")} /></Field>
            </div>

            <SectionTitle icon={Mountain}>Ígnea — 2. Aspectos gerais</SectionTitle>
            <div style={S.formGrid3}>
              <Field label="Cor (alterações, se houver)"><input style={S.input} value={f.corAlteracao} onChange={set("corAlteracao")} placeholder="Ex: esverdeada por alteração" /></Field>
            </div>

            <div style={{ marginTop: 14, fontSize: 13.5, fontWeight: 700, color: "var(--txt-5b7768)", textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 8 }}>Estrutura — 1. Estruturas nas rochas vulcânicas</div>
            <div style={S.formGrid2}>
              <Field label="Pequena escala"><select style={S.input} value={f.estVulcPequena} onChange={set("estVulcPequena")}><option value="">Selecione…</option>{ESTRUTURA_VULC_PEQUENA.map((g) => <option key={g}>{g}</option>)}</select></Field>
              <Field label="Escala"><select style={S.input} value={f.estVulcGrande} onChange={set("estVulcGrande")}><option value="">Selecione…</option>{ESTRUTURA_VULC_GRANDE.map((g) => <option key={g}>{g}</option>)}</select></Field>
            </div>

            <div style={{ marginTop: 14, fontSize: 13.5, fontWeight: 700, color: "var(--txt-5b7768)", textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 8 }}>Estrutura — 2. Estruturas nas rochas plutônicas</div>
            <div style={S.formGrid2}>
              <Field label="Pequena escala"><select style={S.input} value={f.estPlutPequena} onChange={set("estPlutPequena")}><option value="">Selecione…</option>{ESTRUTURA_PLUT_PEQUENA.map((g) => <option key={g}>{g}</option>)}</select></Field>
              <Field label="Escala"><select style={S.input} value={f.estPlutMedia} onChange={set("estPlutMedia")}><option value="">Selecione…</option>{ESTRUTURA_PLUT_MEDIA.map((g) => <option key={g}>{g}</option>)}</select></Field>
            </div>

            
            <div style={{ marginTop: 14, fontSize: 13.5, fontWeight: 700, color: "var(--txt-5b7768)", textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 8 }}>Mineralogia — Félsicos (%)</div>
            <div style={S.formGrid4}>
              <Field label="Feldspato potássico"><input style={S.input} type="number" min="0" max="100" value={f.felFeldspatoK} onChange={set("felFeldspatoK")} placeholder="%" /></Field>
              <Field label="Plagioclásio"><input style={S.input} type="number" min="0" max="100" value={f.felPlagioclasio} onChange={set("felPlagioclasio")} placeholder="%" /></Field>
              <Field label="Quartzo"><input style={S.input} type="number" min="0" max="100" value={f.felQuartzo} onChange={set("felQuartzo")} placeholder="%" /></Field>
              <Field label="Muscovita"><input style={S.input} type="number" min="0" max="100" value={f.felMuscovita} onChange={set("felMuscovita")} placeholder="%" /></Field>
              <Field label="Feldspatoides"><input style={S.input} type="number" min="0" max="100" value={f.felFeldspatoide} onChange={set("felFeldspatoide")} placeholder="%" /></Field>
              <Field label="Outro — mineral"><input style={S.input} value={f.felOutroNome} onChange={set("felOutroNome")} placeholder="Nome do mineral" /></Field>
              <Field label="Outro — %"><input style={S.input} type="number" min="0" max="100" value={f.felOutroPct} onChange={set("felOutroPct")} placeholder="%" /></Field>
            </div>

            <div style={{ marginTop: 14, fontSize: 13.5, fontWeight: 700, color: "var(--txt-5b7768)", textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 8 }}>Mineralogia — Máficos (%)</div>
            <div style={S.formGrid4}>
              <Field label="Turmalina"><input style={S.input} type="number" min="0" max="100" value={f.mafTurmalina} onChange={set("mafTurmalina")} placeholder="%" /></Field>
              <Field label="Anfibólio"><input style={S.input} type="number" min="0" max="100" value={f.mafAnfibolio} onChange={set("mafAnfibolio")} placeholder="%" /></Field>
              <Field label="Biotita"><input style={S.input} type="number" min="0" max="100" value={f.mafBiotita} onChange={set("mafBiotita")} placeholder="%" /></Field>
              <Field label="Magnetita"><input style={S.input} type="number" min="0" max="100" value={f.mafMagnetita} onChange={set("mafMagnetita")} placeholder="%" /></Field>
              <Field label="Outro — mineral"><input style={S.input} value={f.mafOutroNome} onChange={set("mafOutroNome")} placeholder="Nome do mineral" /></Field>
              <Field label="Outro — %"><input style={S.input} type="number" min="0" max="100" value={f.mafOutroPct} onChange={set("mafOutroPct")} placeholder="%" /></Field>
            </div>

            <SomaPct f={f} keys={["felFeldspatoK", "felPlagioclasio", "felQuartzo", "felMuscovita", "felFeldspatoide", "felOutroPct", "mafTurmalina", "mafAnfibolio", "mafBiotita", "mafMagnetita", "mafOutroPct"]} />
            <div style={{ marginTop: 14 }}>
              <Field label="Minerais acessórios"><input style={S.input} value={f.mineralogiaSecundaria} onChange={set("mineralogiaSecundaria")} placeholder="Ex: zircão, apatita, titanita…" /></Field>
            </div>

            <SectionTitle icon={Mountain}>Ígnea — 3. Textura</SectionTitle>
            <div style={S.formGrid3}>
              <Field label="1. Grau de cristalinidade"><select style={S.input} value={f.grauCristalinidade} onChange={set("grauCristalinidade")}><option value="">Selecione…</option>{GRAU_CRISTALINIDADE.map((g) => <option key={g}>{g}</option>)}</select></Field>
              <Field label="2. Grau de visibilidade"><select style={S.input} value={f.grauVisibilidade} onChange={(e) => setF({ ...f, grauVisibilidade: e.target.value, visibilidadeTipo: "" })}><option value="">Selecione…</option>{GRAU_VISIBILIDADE.map((g) => <option key={g}>{g}</option>)}</select></Field>
              {VISIBILIDADE_TIPOS[f.grauVisibilidade] && (
                <Field label={`Tipo (${f.grauVisibilidade.toLowerCase()})`}><select style={S.input} value={f.visibilidadeTipo} onChange={set("visibilidadeTipo")}><option value="">Selecione…</option>{VISIBILIDADE_TIPOS[f.grauVisibilidade].map((g) => <option key={g}>{g}</option>)}</select></Field>
              )}
              <Field label="3. Granulação"><select style={S.input} value={f.tamanhoCristaisAbsoluto} onChange={set("tamanhoCristaisAbsoluto")}><option value="">Selecione…</option>{GRANULACAO_IGNEA.map((g) => <option key={g}>{g}</option>)}</select></Field>
              <Field label="4. Forma geométrica dos cristais"><select style={S.input} value={f.formaGeometricaCristais} onChange={set("formaGeometricaCristais")}><option value="">Selecione…</option>{FORMA_GEOMETRICA_CRISTAIS.map((g) => <option key={g}>{g}</option>)}</select></Field>
            </div>

            <SectionTitle icon={Mountain}>Ígnea — 4. Classificação e nomenclatura</SectionTitle>
            <div style={S.formGrid3}>
              <Field label="1. Teor de sílica"><select style={S.input} value={f.saturacaoSilica} onChange={set("saturacaoSilica")}><option value="">Selecione…</option>{SATURACAO_SILICA.map((g) => <option key={g}>{g}</option>)}</select></Field>
              <Field label="2. Índice de cor (M)"><select style={S.input} value={f.indiceCorIgnea} onChange={set("indiceCorIgnea")}><option value="">Selecione…</option>{INDICE_COR_IGNEA.map((g) => <option key={g}>{g}</option>)}</select></Field>
              <Field label="3. Ambiente de formação e modo de ocorrência"><select style={S.input} value={f.ambienteFormacaoIgnea} onChange={set("ambienteFormacaoIgnea")}><option value="">Selecione…</option>{AMBIENTE_FORMACAO_IGNEA_OPCOES.map((g) => <option key={g}>{g}</option>)}</select></Field>
            </div>

            <div style={{ marginTop: 14 }}><QAPFPanel f={f} setF={setF} /></div>

            <div style={{ marginTop: 14 }}>
              <Field label="Nomenclatura (nome da rocha, segundo o diagrama QAP)"><input style={S.input} value={f.nomenclaturaIgnea} onChange={set("nomenclaturaIgnea")} placeholder="Ex: Granito, Basalto, Andesito…" /></Field>
            </div>
          </>
        )}

        {f.tipoRocha === "Sedimentar" && (
          <>
            <SectionTitle icon={Mountain}>Sedimentar — 1. Localização</SectionTitle>
            <div style={S.formGrid3}>
              {txt("Coordenada X", "coordenadaX")}
              {txt("Coordenada Y", "coordenadaY")}
              {txt("Tipo de afloramento", "tipoAfloramento")}
            </div>

            <SectionTitle icon={Mountain}>Sedimentar — 2. Aspectos gerais</SectionTitle>
            <div style={S.formGrid3}>
              {txt("Cor da rocha", "sedCor", "Ex: cinza-claro, avermelhada")}
              {sel("Origem da cor", "sedCorOrigem", COR_ORIGEM_SED)}
              {sel("Coesão (resistência)", "sedCoesao", SED_COESAO)}
              {sel("Porosidade aparente", "sedPorosidade", SED_POROSIDADE)}
              {sel("Reação ao HCl (ácido diluído)", "rochaReacaoHCl", REACAO_HCL_ROCHA)}
            </div>

            <SectionTitle icon={Mountain}>Sedimentar — 3. Textura</SectionTitle>
            <div style={S.formGrid3}>
              {sel("1. Material de formação", "sedMaterialFormacao", MATERIAL_FORMACAO_SED)}
              {clastica ? (
                <>
                  {sel("2. Granulometria (tamanho dominante dos grãos)", "sedGranulometria", GRANULOMETRIA_SED)}
                  {sel("3. Selecionamento dos grãos", "sedSelecionamento", SELECIONAMENTO_SED)}
                  {sel("4. Arredondamento dos grãos", "sedArredondamento", ARREDONDAMENTO_SED)}
                  {sel("5. Esfericidade dos grãos", "sedEsfericidade", SED_ESFERICIDADE)}
                  {sel("6. Arcabouço", "sedArcabouco", SED_ARCABOUCO)}
                  {sel("7. Cimento", "sedCimento", SED_CIMENTO)}
                </>
              ) : (
                <>
                  {sel("2. Textura (não clástica)", "sedTexturaNaoClastica", SED_TEXTURA_NAO_CLASTICA)}
                  {sel("3. Cimento / material de ligação", "sedCimento", SED_CIMENTO)}
                </>
              )}
            </div>

            <SectionTitle icon={Mountain}>Sedimentar — 4. Estruturas</SectionTitle>
            <div style={S.formGrid3}>
              {sel("1. Estrutura interna (acamamento)", "sedAcamamento", SED_ACAMAMENTO)}
              {sel("2. Espessura das camadas", "sedEspessura", SED_ESPESSURA)}
              {sel("3. Estruturas de superfície", "sedEstrSuperficie", SED_ESTR_SUPERFICIE)}
              {sel("4. Fósseis", "sedFosseis", SED_FOSSEIS)}
              <Field label="5. Outras estruturas (orientação de grãos etc.)?"><select style={S.input} value={f.sedPossuiEstrutura} onChange={set("sedPossuiEstrutura")}><option value="">Selecione…</option><option>Sim</option><option>Não</option></select></Field>
            </div>
            {f.sedPossuiEstrutura === "Sim" && (
              <div style={{ marginTop: 14 }}>
                <Field label="Descreva a estrutura observada"><textarea style={{ ...S.input, minHeight: 60, height: "auto", padding: "11px 13px" }} value={f.sedEstruturaDescricao} onChange={set("sedEstruturaDescricao")} placeholder="Ex: grãos orientados, nódulos, camadas deformadas…" /></Field>
              </div>
            )}

            
            <SectionTitle icon={Mountain}>Sedimentar — 5. Composição mineralógica</SectionTitle>
            <SubTitle>Constituintes principais (%)</SubTitle>
            <div style={S.formGrid4}>
              {pct("Quartzo", "sedQuartzo")}
              {pct("Feldspato", "sedFeldspato")}
              {pct("Fragmentos de rocha (líticos)", "sedLiticos")}
              {pct("Micas", "sedMicas")}
              {pct("Argilominerais", "sedArgilominerais")}
              {pct("Calcita", "sedCalcita")}
              {pct("Dolomita", "sedDolomita")}
              {pct("Óxidos / hidróxidos de ferro", "sedOxidosFe")}
              {pct("Matéria orgânica", "sedMateriaOrganica")}
              {txt("Outro — mineral", "sedOutroNome", "Nome do mineral")}
              {pct("Outro — %", "sedOutroPct")}
            </div>
            <SomaPct f={f} keys={SED_PCT} />
            <div style={{ marginTop: 14 }}>
              {txt("Minerais acessórios", "mineralogiaSecundaria", "Ex: zircão, turmalina, muscovita…")}
            </div>

            <SectionTitle icon={Mountain}>Sedimentar — 6. Classificação e nomenclatura</SectionTitle>
            <div style={S.formGrid3}>
              {sel("1. Grupo da rocha", "sedClassificacao", SED_CLASSIFICACAO)}
              {sel("2. Ambiente deposicional", "sedAmbienteDep", SED_AMBIENTE_DEP)}
              {txt("3. Nomenclatura (nome completo)", "sedNomenclatura", "Ex: Arenito quartzoso, Calcário oolítico…")}
            </div>
            <SugestaoNome sugestao={sugerirNomeSed(f)} atual={f.sedNomenclatura} onUse={(v) => setF({ ...f, sedNomenclatura: v, sedClassificacao: f.sedClassificacao || (SED_CLASSIFICACAO.find((c) => v.startsWith(c)) || "") })} />
            <div style={{ marginTop: 14 }}>
              <PhotoPicker label="7. Desenho / esboço da amostra" value={f.sedEsboco} onChange={(v) => setF({ ...f, sedEsboco: v })} />
            </div>
          </>
        )}

        {f.tipoRocha === "Metamórfica" && (
          <>
            <SectionTitle icon={Mountain}>Metamórfica — 1. Localização</SectionTitle>
            <div style={S.formGrid3}>
              {txt("Coordenada X", "coordenadaX")}
              {txt("Coordenada Y", "coordenadaY")}
              {txt("Tipo de afloramento", "tipoAfloramento")}
            </div>

            <SectionTitle icon={Mountain}>Metamórfica — 2. Aspectos gerais</SectionTitle>
            <div style={S.formGrid3}>
              {txt("Cor da rocha", "metCor", "Ex: cinza-esverdeada, prateada")}
              {sel("Brilho", "metBrilho", MET_BRILHO)}
              {sel("Reação ao HCl (ácido diluído)", "rochaReacaoHCl", REACAO_HCL_ROCHA)}
            </div>

            <SectionTitle icon={Mountain}>Metamórfica — 3. Estrutura</SectionTitle>
            <div style={S.formGrid3}>
              {sel("1. Aspecto — a rocha parece", "metAspecto", MET_ASPECTO)}
              {sel("2. Estruturas associadas", "metEstrAssoc", MET_ESTRUTURAS_ASSOC)}
            </div>

            <SectionTitle icon={Mountain}>Metamórfica — 4. Textura</SectionTitle>
            <div style={S.formGrid3}>
              {sel("1. Granulação — os grãos da rocha são", "metGranulacao", MET_GRANULACAO)}
              {sel("2. Textura", "metTextura", MET_TEXTURA)}
              {sel("3. Forma dos cristais", "metForma", MET_FORMA_CRISTAIS)}
              {sel("4. Porfiroblastos", "metPorfiroblastos", MET_PORFIROBLASTOS)}
            </div>

            
            <SectionTitle icon={Mountain}>Metamórfica — 5. Mineralogia</SectionTitle>
            <SubTitle>Constituintes principais (%)</SubTitle>
            <div style={S.formGrid4}>
              {pct("Quartzo", "metQuartzo")}
              {pct("Feldspato", "metFeldspato")}
              {pct("Carbonatos (calcita / dolomita)", "metCarbonato")}
              {pct("Muscovita", "metMuscovita")}
              {pct("Biotita", "metBiotita")}
              {pct("Clorita", "metClorita")}
              {pct("Anfibólio", "metAnfibolio")}
              {pct("Piroxênio", "metPiroxenio")}
              {pct("Granada", "metGranada")}
              {pct("Talco / serpentina", "metTalcoSerp")}
              {pct("Grafita", "metGrafita")}
              {pct("Aluminossilicatos (cianita, silimanita, andaluzita)", "metAluminossilicato")}
              {txt("Outro — mineral", "metOutroNome", "Nome do mineral")}
              {pct("Outro — %", "metOutroPct")}
            </div>
            <SomaPct f={f} keys={MET_PCT} />
            <div style={{ marginTop: 14 }}>
              <Field label="Minerais-índice e outros minerais observados"><textarea style={{ ...S.input, minHeight: 60, height: "auto", padding: "11px 13px" }} value={f.mineraisIndice} onChange={set("mineraisIndice")} placeholder="Ex: muscovita, quartzo, granada…" /></Field>
            </div>
            <div style={{ marginTop: 14 }}>
              {txt("Minerais acessórios", "mineralogiaSecundaria", "Ex: zircão, apatita, turmalina, rutilo…")}
            </div>

            <SectionTitle icon={Mountain}>Metamórfica — 6. Classificação e nomenclatura</SectionTitle>
            <div style={S.formGrid3}>
              {sel("1. Tipo de metamorfismo", "tipoMetamorfismo", TIPO_METAMORFISMO)}
              {sel("2. Grau metamórfico", "grauMetamorfico", GRAU_METAMORFICO)}
              {sel("3. Fácies metamórfica", "metFacies", MET_FACIES)}
              {sel("4. Protólito — lembra alguma “rocha mãe”?", "rochaOriginal", MET_PROTOLITO)}
              {txt("5. Nomenclatura (nome da rocha)", "metNomenclatura", "Ex: Xisto, Gnaisse, Mármore…")}
            </div>
            <SugestaoNome sugestao={sugerirNomeMet(f)} atual={f.metNomenclatura} onUse={(v) => setF({ ...f, metNomenclatura: v })} />
          </>
        )}

        </div>
        <div className="geoWizardPanel" style={{ display: step === 2 ? "block" : "none" }}>
        <SectionTitle icon={ClipboardList}>Registro</SectionTitle>
        <div style={S.formGrid3}>
          <Field label="Data do cadastro"><input type="date" style={S.input} value={f.dataCadastro} onChange={set("dataCadastro")} /></Field>
          <Field label="Responsável"><input style={S.input} value={f.responsavel} onChange={set("responsavel")} /></Field>
          <Field label="Observações" span={3}><textarea style={{ ...S.input, minHeight: 60, height: "auto", padding: "11px 13px" }} value={f.observacoes} onChange={set("observacoes")} /></Field>
        </div>

        <SectionTitle icon={ImageIcon}>Fotos</SectionTitle>
        <PhotoPicker label="Fotos da amostra (até 6)" value={f.fotos} onChange={(v) => setF({ ...f, fotos: v })} multiple />

        </div>
        <WizardActions current={step} total={steps.length}
          onBack={() => setStep((v) => Math.max(0, v - 1))}
          onNext={() => setStep((v) => Math.min(steps.length - 1, v + 1))}
          onCancel={onClose} onSave={submit} saveLabel="Salvar rocha" />
      </div>
    </ModalShell>
  );
}

// Geometria do losango (compartilhada pelo painel pequeno e pelo QAPF completo)
const QG = { cx: 190, yQ: 26, yAP: 236, yF: 446, hw: 166 };
const qgPos = (lado, nivel, t) => {
  const lv = nivel / 100;
  const y = lado === "Q" ? QG.yAP - (QG.yAP - QG.yQ) * lv : QG.yAP + (QG.yF - QG.yAP) * lv;
  return [QG.cx + (t / 100 - 0.5) * 2 * QG.hw * (1 - lv), y];
};

function QAPFSvg({ modo, res, numeros, hoverId, onHover, onPick }) {
  const campos = QAPF_CAMPOS[modo] || [];
  const ativo = res ? findCampoQAPF(modo, res.lado, res.nivel, res.t) : null;
  const pt = res ? qgPos(res.lado, res.nivel, res.t) : null;
  const pick = (e) => {
    if (!onPick) return;
    const svg = e.currentTarget;
    const r = svg.getBoundingClientRect();
    const x = ((e.clientX - r.left) / r.width) * 380, y = ((e.clientY - r.top) / r.height) * 472;
    const lado = y <= QG.yAP ? "Q" : "F";
    const alt = lado === "Q" ? QG.yAP - QG.yQ : QG.yF - QG.yAP;
    const nivel = Math.max(0, Math.min(100, (Math.abs(y - QG.yAP) / alt) * 100));
    const larg = 2 * QG.hw * (1 - nivel / 100);
    const t = larg < 0.5 ? 50 : Math.max(0, Math.min(100, ((x - QG.cx) / larg + 0.5) * 100));
    onPick({ lado, nivel, t });
  };
  const tick = (key, x, y, n, left) => <text key={key} x={left ? x - 4 : x + 4} y={y + 3} fontSize="8.5" textAnchor={left ? "end" : "start"} fill="var(--txt-5b7768)">{n}</text>;
  const ticks = [];
  [5, 20, 60, 90].forEach((n) => { const [xl, yl] = qgPos("Q", n, 0), [xr, yr] = qgPos("Q", n, 100); ticks.push(tick(`ql${n}`, xl, yl, n, true), tick(`qr${n}`, xr, yr, n, false)); });
  [10, 60, 90].forEach((n) => { if (modo === "Plutônica" && n === 90) return; const [xl, yl] = qgPos("F", n, 0), [xr, yr] = qgPos("F", n, 100); ticks.push(tick(`fl${n}`, xl, yl, n, true), tick(`fr${n}`, xr, yr, n, false)); });
  return (
    <svg
      viewBox="0 0 380 472"
      style={{ width: "100%", height: "auto", touchAction: onPick ? "none" : "auto", cursor: onPick ? "crosshair" : "default" }}
      role="img" aria-label="Diagrama QAPF"
      onPointerDown={onPick ? (e) => { e.currentTarget.setPointerCapture?.(e.pointerId); pick(e); } : undefined}
      onPointerMove={onPick ? (e) => { if (e.buttons === 1) pick(e); } : undefined}
    >
      <polygon points={`${QG.cx},${QG.yQ} ${QG.cx + QG.hw},${QG.yAP} ${QG.cx},${QG.yF} ${QG.cx - QG.hw},${QG.yAP}`} fill="var(--bg-f7f9f6)" />
      {campos.map((c) => {
        const pts = [qgPos(c.lado, c.n1, c.t1), qgPos(c.lado, c.n1, c.t2), qgPos(c.lado, c.n2, c.t2), qgPos(c.lado, c.n2, c.t1)].map((p) => p.join(",")).join(" ");
        const on = ativo && ativo.id === c.id;
        return (
          <polygon
            key={c.id} points={pts}
            fill={on ? "rgba(46,125,91,.32)" : hoverId === c.id ? "rgba(201,162,39,.28)" : "transparent"}
            stroke="var(--bd-0b3d2e)" strokeWidth="1" strokeLinejoin="round"
            onMouseEnter={onHover ? () => onHover(c.id) : undefined}
            onMouseLeave={onHover ? () => onHover(null) : undefined}
          />
        );
      })}
      <line x1={QG.cx - QG.hw} y1={QG.yAP} x2={QG.cx + QG.hw} y2={QG.yAP} stroke="var(--bd-0b3d2e)" strokeWidth="1.8" style={{ pointerEvents: "none" }} />
      <polygon points={`${QG.cx},${QG.yQ} ${QG.cx + QG.hw},${QG.yAP} ${QG.cx},${QG.yF} ${QG.cx - QG.hw},${QG.yAP}`} fill="none" stroke="var(--bd-0b3d2e)" strokeWidth="1.8" style={{ pointerEvents: "none" }} />
      {ticks}
      {numeros && campos.map((c) => {
        const [x, y] = qgPos(c.lado, (c.n1 + c.n2) / 2, (c.t1 + c.t2) / 2);
        return <text key={`n${c.id}`} x={x} y={y + 3} fontSize={c.n2 - c.n1 <= 10 ? 7.5 : 9.5} fontWeight="700" textAnchor="middle" fill="var(--txt-0b3d2e)" style={{ pointerEvents: "none" }}>{c.id}</text>;
      })}
      <text x={QG.cx} y={QG.yQ - 8} fontSize="13" fontWeight="700" textAnchor="middle" fill="var(--txt-0b3d2e)">Q</text>
      <text x={QG.cx - QG.hw - 12} y={QG.yAP + 4} fontSize="13" fontWeight="700" textAnchor="end" fill="var(--txt-0b3d2e)">A</text>
      <text x={QG.cx + QG.hw + 12} y={QG.yAP + 4} fontSize="13" fontWeight="700" fill="var(--txt-0b3d2e)">P</text>
      <text x={QG.cx} y={QG.yF + 18} fontSize="13" fontWeight="700" textAnchor="middle" fill="var(--txt-0b3d2e)">F</text>
      {pt && (
        <g style={{ pointerEvents: "none" }}>
          <circle cx={pt[0]} cy={pt[1]} r="9" fill="#d1495b" opacity="0.28" />
          <circle cx={pt[0]} cy={pt[1]} r="4.5" fill="#d1495b" stroke="#fff" strokeWidth="1.5" />
        </g>
      )}
    </svg>
  );
}

function QAPFModoToggle({ modo, setModo }) {
  return (
    <div style={{ display: "flex", maxWidth: 320 }}>
      {["Plutônica", "Vulcânica"].map((m) => (
        <button type="button" key={m} onClick={() => setModo(m)} style={{ flex: 1, padding: "8px 10px", border: "1.5px solid " + (modo === m ? "#2E7D5B" : "var(--bd-dfe8e2)"), background: modo === m ? "#2E7D5B" : "var(--bg-fff)", color: modo === m ? "#fff" : "var(--txt-5b7768)", fontWeight: 700, fontSize: 14.5, borderRadius: m === "Plutônica" ? "9px 0 0 9px" : "0 9px 9px 0" }}>{m === "Plutônica" ? "Plutônicas" : "Vulcânicas"}</button>
      ))}
    </div>
  );
}

// QAPF completo: tela grande, interativa (clique/arraste no diagrama, digite Q-A-P-F, passe o mouse nos campos)
function QAPFModal({ f, setF, modoInicial, onClose }) {
  const [modo, setModo] = useState(modoInicial);
  const fmt = (v) => { const n = num(v); return n ? String(Math.round(n * 10) / 10) : ""; };
  const [v, setV] = useState({ q: fmt(f.felQuartzo), a: fmt(f.felFeldspatoK), p: fmt(f.felPlagioclasio), f: fmt(f.felFeldspatoide) });
  const [hoverId, setHoverId] = useState(null);
  const res = classificarQAPF(v.q, v.a, v.p, v.f, modo);
  const nome = nomeSugeridoQAPF(res, f.saturacaoSilica);
  const campos = QAPF_CAMPOS[modo];
  const hoverCampo = campos.find((c) => c.id === hoverId);
  const ativo = res ? findCampoQAPF(modo, res.lado, res.nivel, res.t) : null;
  const round1 = (x) => String(Math.round(x * 10) / 10);
  const setPonto = ({ lado, nivel, t }) => {
    const resto = 100 - nivel;
    const A = resto * (1 - t / 100), P = resto * (t / 100);
    setV(lado === "Q" ? { q: round1(nivel), a: round1(A), p: round1(P), f: "" } : { q: "", a: round1(A), p: round1(P), f: round1(nivel) });
  };
  const centro = (c) => setPonto({ lado: c.lado, nivel: Math.min(99, (c.n1 + c.n2) / 2), t: (c.t1 + c.t2) / 2 });
  const aplicar = (comValores) => {
    const base = { ...f, nomenclaturaIgnea: nome || f.nomenclaturaIgnea };
    if (comValores && res) Object.assign(base, { felQuartzo: res.Q ? round1(res.Q) : "", felFeldspatoK: round1(res.A), felPlagioclasio: round1(res.P), felFeldspatoide: res.F ? round1(res.F) : "" });
    setF(base); onClose();
  };
  const inp = (k, lab) => (
    <Field label={lab}><input style={S.input} type="number" min="0" max="100" step="0.1" value={v[k]} onChange={(e) => setV({ ...v, [k]: e.target.value })} placeholder="%" /></Field>
  );
  return (
    <div style={{ ...S.modalOverlay, zIndex: 70 }} onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div style={{ ...S.modal, maxWidth: 1080 }}>
        <div style={S.modalHead}>
          <h3 style={{ fontFamily: "Fraunces", fontSize: 20, color: "var(--txt-0b3d2e)", margin: 0 }}>Diagrama QAPF completo</h3>
          <button type="button" style={S.iconBtn} onClick={onClose}><X size={17} /></button>
        </div>
        <div style={{ ...S.modalBody, paddingTop: 16 }}>
          <div style={{ fontSize: 14.5, color: "var(--txt-5b7768)", marginBottom: 12 }}>
            Clique ou arraste dentro do losango para mover o ponto, digite as porcentagens ou escolha um campo na lista. Passe o mouse sobre um campo para ver o nome.
          </div>
          <div style={{ marginBottom: 14 }}><QAPFModoToggle modo={modo} setModo={setModo} /></div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 22, alignItems: "flex-start" }}>
            <div style={{ flex: "1 1 460px", maxWidth: 640, minWidth: 280 }}>
              <QAPFSvg modo={modo} res={res} numeros hoverId={hoverId} onHover={setHoverId} onPick={setPonto} />
            </div>
            <div style={{ flex: "1 1 320px", minWidth: 280 }}>
              <div style={{ ...S.formGrid4, marginTop: 0 }}>
                {inp("q", "Q — quartzo")}{inp("a", "A — feld. alcalino")}{inp("p", "P — plagioclásio")}{inp("f", "F — feldspatoide")}
              </div>
              <div style={{ background: "var(--bg-f2f7f4)", border: "1px solid var(--bd-eef3f0)", borderRadius: 12, padding: 14, margin: "14px 0" }}>
                <div style={{ fontSize: 13.5, fontWeight: 700, color: "var(--txt-5b7768)", textTransform: "uppercase", letterSpacing: 0.4 }}>
                  {hoverCampo ? `Campo ${hoverCampo.id} (passando o mouse)` : res ? `Campo ${ativo ? ativo.id : ""} — classificação` : "Classificação"}
                </div>
                <div style={{ fontFamily: "Fraunces", fontSize: 22, fontWeight: 600, color: "var(--txt-0b3d2e)", lineHeight: 1.25, margin: "4px 0" }}>
                  {hoverCampo ? hoverCampo.nome : res ? nome : "Informe Q, A, P e F ou clique no diagrama"}
                </div>
                {res && !hoverCampo && (
                  <div style={{ fontSize: 14, color: "var(--txt-5b7768)" }}>
                    Q {res.Q.toFixed(1)}% · A {res.A.toFixed(1)}% · P {res.P.toFixed(1)}% · F {res.F.toFixed(1)}% · P/(A+P) = {res.t.toFixed(0)}%
                  </div>
                )}
                {res && res.aviso && !hoverCampo && <div style={{ ...S.tipoHint, marginTop: 8 }}><AlertCircle size={14} /> {res.aviso}</div>}
              </div>
              <div style={{ fontSize: 13.5, fontWeight: 700, color: "var(--txt-5b7768)", textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 6 }}>Campos do diagrama ({modo === "Plutônica" ? "plutônicas" : "vulcânicas"})</div>
              <div style={{ maxHeight: 300, overflowY: "auto", border: "1px solid var(--bd-eef3f0)", borderRadius: 10 }}>
                {campos.map((c) => (
                  <div
                    key={c.id}
                    onClick={() => centro(c)}
                    onMouseEnter={() => setHoverId(c.id)}
                    onMouseLeave={() => setHoverId(null)}
                    style={{ display: "flex", gap: 10, padding: "7px 10px", fontSize: 14.5, cursor: "pointer", color: "var(--txt-1c2a22)", background: ativo && ativo.id === c.id ? "var(--bg-e4f2e9)" : hoverId === c.id ? "var(--bg-f2f7f4)" : "transparent", borderBottom: "1px solid var(--bd-eef3f0)" }}
                  >
                    <b style={{ minWidth: 20, color: "var(--txt-2e7d5b)" }}>{c.id}</b><span>{c.nome}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
          <div style={{ fontSize: 13, color: "var(--txt-7c9186)", marginTop: 14 }}>
            {modo === "Plutônica" ? "Diagrama QAPF para classificação modal de rochas plutônicas (baseado em Streckeisen, 1976)." : "Diagrama QAPF para classificação modal de rochas vulcânicas (baseado em Streckeisen, 1978)."} Q = quartzo, A = feldspato alcalino, P = plagioclásio, F = feldspatoides. Não usar quando M (minerais máficos) &gt; 90%.
          </div>
          <div style={S.modalFooter}>
            <button type="button" style={S.btnGhost} onClick={onClose}>Fechar</button>
            <button type="button" style={S.btnGhost} disabled={!res} onClick={() => aplicar(false)}>Aplicar só a nomenclatura</button>
            <button type="button" style={{ ...S.btnPrimary, width: "auto", marginTop: 0, padding: "11px 22px", opacity: res ? 1 : 0.5 }} disabled={!res} onClick={() => aplicar(true)}>Aplicar valores e nomenclatura</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function QAPFPanel({ f, setF }) {
  const modoForm = (f.ambienteFormacaoIgnea || "").startsWith("Vulcânica") ? "Vulcânica" : (f.ambienteFormacaoIgnea || "").startsWith("Plutônica") ? "Plutônica" : "";
  const [modo, setModo] = useState(modoForm || "Plutônica");
  const [aberto, setAberto] = useState(false);
  useEffect(() => { if (modoForm) setModo(modoForm); }, [modoForm]);
  const res = classificarQAPF(f.felQuartzo, f.felFeldspatoK, f.felPlagioclasio, f.felFeldspatoide, modo);
  const maficos = num(f.mafTurmalina) + num(f.mafAnfibolio) + num(f.mafBiotita) + num(f.mafMagnetita) + num(f.mafOutroPct);
  const nomeSug = nomeSugeridoQAPF(res, f.saturacaoSilica);
  return (
    <div style={{ border: "1px solid var(--bd-dfe8e2)", borderRadius: 12, padding: 16, background: "var(--bg-fff)" }}>
      {aberto && <QAPFModal f={f} setF={setF} modoInicial={modo} onClose={() => setAberto(false)} />}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap", marginBottom: 4 }}>
        <div style={{ fontSize: 15, fontWeight: 700, color: "var(--txt-0b3d2e)" }}>Diagrama QAPF — classificação modal</div>
        <button type="button" style={{ ...S.btnGhost, padding: "7px 14px", fontSize: 14.5 }} onClick={() => setAberto(true)}>Abrir QAPF completo</button>
      </div>
      <div style={{ fontSize: 14, color: "var(--txt-5b7768)", marginBottom: 12 }}>
        Usa as porcentagens de <b>Quartzo (Q)</b>, <b>Feldspato potássico (A)</b>, <b>Plagioclásio (P)</b> e <b>Feldspatoides (F)</b> da mineralogia félsica. Os valores são reescalonados para somar 100%.
      </div>
      <div style={{ marginBottom: 12 }}><QAPFModoToggle modo={modo} setModo={setModo} /></div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 18, alignItems: "flex-start" }}>
        <div style={{ flex: "0 1 380px", width: "100%", maxWidth: 380 }}><QAPFSvg modo={modo} res={res} /></div>
        <div style={{ flex: "1 1 240px", minWidth: 220 }}>
          {!res && <div style={{ fontSize: 15, color: "var(--txt-5b7768)", background: "var(--bg-f2f7f4)", borderRadius: 10, padding: 14 }}>Preencha as porcentagens de quartzo, feldspato potássico, plagioclásio e/ou feldspatoides na mineralogia félsica para ver o ponto no diagrama e a classificação. Ou use <b>Abrir QAPF completo</b> para clicar direto no diagrama.</div>}
          {res && (
            <>
              <div style={{ fontSize: 13.5, fontWeight: 700, color: "var(--txt-5b7768)", textTransform: "uppercase", letterSpacing: 0.4 }}>Classificação sugerida ({modo === "Plutônica" ? "plutônica" : "vulcânica"})</div>
              <div style={{ fontFamily: "Fraunces", fontSize: 21, fontWeight: 600, color: "var(--txt-0b3d2e)", margin: "4px 0 12px", lineHeight: 1.25 }}>{nomeSug}</div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 8, marginBottom: 12 }}>
                {[["Q", res.Q], ["A", res.A], ["P", res.P], ["F", res.F]].map(([k, val]) => (
                  <div key={k} style={{ background: "var(--bg-f2f7f4)", borderRadius: 9, padding: "8px 6px", textAlign: "center" }}>
                    <div style={{ fontSize: 13, fontWeight: 700, color: "var(--txt-5b7768)" }}>{k}</div>
                    <div style={{ fontSize: 16, fontWeight: 700, color: "var(--txt-0b3d2e)" }}>{val.toFixed(1)}%</div>
                  </div>
                ))}
              </div>
              <div style={{ fontSize: 14, color: "var(--txt-5b7768)", marginBottom: 10 }}>Razão P/(A+P) = <b>{res.t.toFixed(0)}%</b></div>
              {res.aviso && <div style={S.tipoHint}><AlertCircle size={14} /> {res.aviso}</div>}
              {maficos > 90 && <div style={{ ...S.tipoHint, marginTop: 8 }}><AlertCircle size={14} /> Minerais máficos somam {maficos.toFixed(0)}% (M &gt; 90). O diagrama QAPF não deve ser usado nesse caso (rocha ultramáfica).</div>}
              {nomeSug.includes("/") && <div style={{ fontSize: 13.5, color: "var(--txt-5b7768)", marginTop: 8 }}>Mais de um nome possível neste campo. Diferencie pelo teor de sílica, pelo plagioclásio (anortita) ou pelo índice de cor.</div>}
              <button type="button" style={{ ...S.btnGhost, marginTop: 12 }} onClick={() => setF({ ...f, nomenclaturaIgnea: nomeSug })}>Usar como nomenclatura</button>
            </>
          )}
        </div>
      </div>
      <div style={{ fontSize: 13, color: "var(--txt-7c9186)", marginTop: 12 }}>
        {modo === "Plutônica" ? "Diagrama QAPF para classificação modal de rochas plutônicas (baseado em Streckeisen, 1976)." : "Diagrama QAPF para classificação modal de rochas vulcânicas (baseado em Streckeisen, 1978)."} Não usar quando M (minerais máficos) &gt; 90%.
      </div>
    </div>
  );
}

function SectionTitle({ icon: Icon, children }) {
  return <div style={S.sectionTitle}><span style={S.sectionIcon}><Icon size={15} /></span><span>{children}</span></div>;
}

function FichaModal({ item, onClose, onEdit, onExportXML }) {
  const isMin = item.tipoAmostra === "mineral";
  const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=180x180&margin=6&data=${encodeURIComponent(item.codigo)}`;
  const fotos = isMin ? [item.fotoPrincipal, ...(item.fotosAdicionais || [])].filter(Boolean) : (item.fotos || []);
  return (
    <div style={S.modalOverlay} onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }} className="printArea">
      <div style={{ ...S.modal, maxWidth: 760 }} id="fichaPrint">
        <div style={S.modalHead} className="noPrint">
          <h3 style={{ fontFamily: "Fraunces", fontSize: 20, color: "var(--txt-0b3d2e)", margin: 0 }}>Ficha da Amostra</h3>
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
              <div style={{ fontFamily: "Fraunces", fontWeight: 600, fontSize: 25, color: "var(--txt-0b3d2e)" }}>{item.nome}</div>
              <div style={{ color: "var(--txt-5b7768)", fontSize: 15.5 }}>{isMin ? item.grupo : item.tipoRocha}</div>
              <span style={S.mono}>{item.codigo}</span>
            </div>
            <img src={qrUrl} alt="QR Code" style={{ width: 84, height: 84 }} />
          </div>

          {fotos.length > 0 && (
            <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" }}>
              {fotos.map((f, i) => <img key={i} src={f} style={{ width: 100, height: 100, objectFit: "cover", borderRadius: 8, border: "1px solid var(--bd-dfe8e2)" }} />)}
            </div>
          )}

          {isMin ? (
            <>
              <FichaGrupo titulo="Identificação" campos={[["Grupo mineral", item.grupo], ["Composição química", item.composicaoQuimica]]} />
              <FichaGrupo titulo="Características físicas" campos={[["Cor", item.cor], ["Traço", item.traco], ["Brilho", [item.brilho, item.brilhoDetalhe].filter(Boolean).join(" — ")], ["Diafaneidade", item.transparencia], ["Clivagem", item.clivagem], ["Quantidade de planos", item.planosClivagem], ["Fratura", item.fratura], ["Hábito", item.habito], ["Sistema cristalino", item.sistemaCristalino], ["Dureza (Mohs)", item.dureza], ["Densidade", item.densidade], ["Magnetismo", item.magnetismo], ["Fluorescência", item.fluorescencia], ["Tenacidade", item.tenacidade], ["Reação com ácido", item.reacaoAcido]]} />
              <FichaGrupo titulo="Complementares" campos={[["Ambiente de formação", item.ambienteFormacao], ["Usos", item.usos], ["Conservação", item.estadoConservacao]]} />
            </>
          ) : (
            <>
              {item.tipoRocha === "Ígnea" && (
                <>
                  <FichaGrupo titulo="Ígnea — Localização" campos={[["Coordenada X", item.coordenadaX], ["Coordenada Y", item.coordenadaY], ["Tipo de afloramento", item.tipoAfloramento]]} />
                  <FichaGrupo titulo="Ígnea — Aspectos gerais" campos={[["Cor (alterações)", item.corAlteracao]]} />
                  <FichaGrupo titulo="Ígnea — Estruturas vulcânicas" campos={[["Pequena escala", item.estVulcPequena], ["Escala", item.estVulcGrande]]} />
                  <FichaGrupo titulo="Ígnea — Estruturas plutônicas" campos={[["Pequena escala", item.estPlutPequena], ["Escala", item.estPlutMedia]]} />
                  <FichaGrupo titulo="Ígnea — Mineralogia félsica (%)" campos={[["Feldspato potássico", item.felFeldspatoK], ["Plagioclásio", item.felPlagioclasio], ["Quartzo", item.felQuartzo], ["Muscovita", item.felMuscovita], ["Feldspatoides", item.felFeldspatoide], ["Outro mineral", item.felOutroNome], ["Outro mineral (%)", item.felOutroPct]]} />
                  <FichaGrupo titulo="Ígnea — Mineralogia máfica (%)" campos={[["Turmalina", item.mafTurmalina], ["Anfibólio", item.mafAnfibolio], ["Biotita", item.mafBiotita], ["Magnetita", item.mafMagnetita], ["Outro mineral", item.mafOutroNome], ["Outro mineral (%)", item.mafOutroPct]]} />
                  <FichaGrupo titulo="Ígnea — Minerais acessórios" campos={[["Minerais acessórios", item.mineralogiaSecundaria]]} />
                  <FichaGrupo titulo="Ígnea — Textura" campos={[["Grau de cristalinidade", item.grauCristalinidade], ["Grau de visibilidade", [item.grauVisibilidade, item.visibilidadeTipo].filter(Boolean).join(" — ")], ["Granulação", item.tamanhoCristaisAbsoluto], ["Forma geométrica", item.formaGeometricaCristais]]} />
                  <FichaGrupo titulo="Ígnea — Classificação" campos={[["Teor de sílica", item.saturacaoSilica], ["Índice de cor (M)", item.indiceCorIgnea], ["Ambiente de formação", item.ambienteFormacaoIgnea], ["Nomenclatura", item.nomenclaturaIgnea]]} />
                  {/* Registros antigos */}
                  <FichaGrupo titulo="Ígnea — campos anteriores" campos={[["Estrutura vulcânica", item.estruturaVulcanica], ["Estrutura plutônica", item.estruturaPlutonica], ["Tamanho relativo", item.tamanhoCristaisRelativo]]} />
                </>
              )}
              {item.tipoRocha === "Sedimentar" && (
                <>
                  <FichaGrupo titulo="Sedimentar — Localização" campos={[["Coordenada X", item.coordenadaX], ["Coordenada Y", item.coordenadaY], ["Tipo de afloramento", item.tipoAfloramento]]} />
                  <FichaGrupo titulo="Sedimentar — Aspectos gerais" campos={[["Cor da rocha", item.sedCor], ["Origem da cor", item.sedCorOrigem], ["Coesão", item.sedCoesao], ["Porosidade aparente", item.sedPorosidade], ["Reação ao HCl", item.rochaReacaoHCl]]} />
                  <FichaGrupo titulo="Sedimentar — Textura" campos={[["Material de formação", item.sedMaterialFormacao], ["Granulometria", item.sedGranulometria], ["Selecionamento", item.sedSelecionamento], ["Arredondamento", item.sedArredondamento], ["Esfericidade", item.sedEsfericidade], ["Arcabouço", item.sedArcabouco], ["Textura (não clástica)", item.sedTexturaNaoClastica], ["Cimento", item.sedCimento]]} />
                  <FichaGrupo titulo="Sedimentar — Estruturas" campos={[["Estrutura interna", item.sedAcamamento], ["Espessura das camadas", item.sedEspessura], ["Estruturas de superfície", item.sedEstrSuperficie], ["Fósseis", item.sedFosseis], ["Outras estruturas", [item.sedPossuiEstrutura, item.sedEstruturaDescricao].filter(Boolean).join(" — ")]]} />
                  <FichaGrupo titulo="Sedimentar — Composição (%)" campos={[["Quartzo", item.sedQuartzo], ["Feldspato", item.sedFeldspato], ["Fragmentos de rocha", item.sedLiticos], ["Micas", item.sedMicas], ["Argilominerais", item.sedArgilominerais], ["Calcita", item.sedCalcita], ["Dolomita", item.sedDolomita], ["Óxidos de ferro", item.sedOxidosFe], ["Matéria orgânica", item.sedMateriaOrganica], ["Outro mineral", item.sedOutroNome], ["Outro mineral (%)", item.sedOutroPct], ["Minerais acessórios", item.mineralogiaSecundaria]]} />
                  <FichaGrupo titulo="Sedimentar — Classificação" campos={[["Grupo da rocha", item.sedClassificacao], ["Ambiente deposicional", item.sedAmbienteDep], ["Nomenclatura", item.sedNomenclatura]]} />
                  {item.sedEsboco && (
                    <div style={{ marginBottom: 16 }}>
                      <div style={S.fichaSecTitulo}>Desenho / esboço da amostra</div>
                      <img src={item.sedEsboco} alt="Esboço da amostra" style={{ maxWidth: 260, width: "100%", borderRadius: 8, border: "1px solid var(--bd-dfe8e2)" }} />
                    </div>
                  )}
                  {/* Registros feitos antes da ficha nova continuam exibindo os campos antigos */}
                  <FichaGrupo titulo="Sedimentar — campos anteriores" campos={[["Tipo sedimentar", item.tipoSedimentar], ["Cimento/matriz", item.cimentoMatriz], ["Selecionamento (antigo)", item.selecionamento], ["Arredondamento (antigo)", item.arredondamento], ["Presença de fósseis", item.presencaFosseis]]} />
                </>
              )}
              {item.tipoRocha === "Metamórfica" && (
                <>
                  <FichaGrupo titulo="Metamórfica — Localização" campos={[["Coordenada X", item.coordenadaX], ["Coordenada Y", item.coordenadaY], ["Tipo de afloramento", item.tipoAfloramento]]} />
                  <FichaGrupo titulo="Metamórfica — Aspectos gerais" campos={[["Cor da rocha", item.metCor], ["Brilho", item.metBrilho], ["Reação ao HCl", item.rochaReacaoHCl]]} />
                  <FichaGrupo titulo="Metamórfica — Estrutura" campos={[["Aspecto", item.metAspecto], ["Estruturas associadas", item.metEstrAssoc]]} />
                  <FichaGrupo titulo="Metamórfica — Textura" campos={[["Granulação", item.metGranulacao], ["Textura", item.metTextura], ["Forma dos cristais", item.metForma], ["Porfiroblastos", item.metPorfiroblastos]]} />
                  <FichaGrupo titulo="Metamórfica — Mineralogia (%)" campos={[["Quartzo", item.metQuartzo], ["Feldspato", item.metFeldspato], ["Carbonatos", item.metCarbonato], ["Muscovita", item.metMuscovita], ["Biotita", item.metBiotita], ["Clorita", item.metClorita], ["Anfibólio", item.metAnfibolio], ["Piroxênio", item.metPiroxenio], ["Granada", item.metGranada], ["Talco / serpentina", item.metTalcoSerp], ["Grafita", item.metGrafita], ["Aluminossilicatos", item.metAluminossilicato], ["Outro mineral", item.metOutroNome], ["Outro mineral (%)", item.metOutroPct], ["Minerais-índice / observados", item.mineraisIndice], ["Minerais acessórios", item.mineralogiaSecundaria]]} />
                  <FichaGrupo titulo="Metamórfica — Classificação" campos={[["Tipo de metamorfismo", item.tipoMetamorfismo], ["Grau metamórfico", item.grauMetamorfico], ["Fácies metamórfica", item.metFacies], ["Protólito", item.rochaOriginal], ["Nomenclatura", item.metNomenclatura]]} />
                  {/* Registros antigos */}
                  <FichaGrupo titulo="Metamórfica — campos anteriores" campos={[["Foliação", item.foliacao]]} />
                </>
              )}
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
:root{--bd-0b3d2e:#0b3d2e;--bd-cfe0d6:#cfe0d6;--bd-dfe8e2:#dfe8e2;--bd-e4f2e9:#e4f2e9;--bd-e7eee9:#e7eee9;--bd-e8dcae:#e8dcae;--bd-eef3f0:#eef3f0;--bd-f2f7f4:#f2f7f4;--bd-fff:#fff;--bg-cfe0d6:#cfe0d6;--bg-e4f2e9:#e4f2e9;--bg-eef3f0:#eef3f0;--bg-efefef:#efefef;--bg-f2f7f4:#f2f7f4;--bg-f7f9f6:#f7f9f6;--bg-fafcf9:#fafcf9;--bg-fbeaea:#fbeaea;--bg-fbf2da:#fbf2da;--bg-fff:#fff;--bg-fffcf2:#fffcf2;--bg-ffffff:#ffffff;--txt-0b3d2e:#0b3d2e;--txt-1c2a22:#1c2a22;--txt-223028:#223028;--txt-2e7d5b:#2e7d5b;--txt-3a5044:#3a5044;--txt-5b7768:#5b7768;--txt-6b5817:#6b5817;--txt-7a5f0a:#7a5f0a;--txt-7c9186:#7c9186;--txt-8a2f2f:#8a2f2f;--txt-9a7b0a:#9a7b0a;--txt-a9bdb2:#a9bdb2}
:root[data-theme="dark"]{--bd-0b3d2e:#8fd0b0;--bd-cfe0d6:#33493e;--bd-dfe8e2:#2f453a;--bd-e4f2e9:#2b3d33;--bd-e7eee9:#2a3f34;--bd-e8dcae:#5a4f22;--bd-eef3f0:#223229;--bd-f2f7f4:#223229;--bd-fff:#17251d;--bg-cfe0d6:#2b4538;--bg-e4f2e9:#1f3a2d;--bg-eef3f0:#1e2f26;--bg-efefef:#2a2f2c;--bg-f2f7f4:#1a2a22;--bg-f7f9f6:#0d1511;--bg-fafcf9:#17251d;--bg-fbeaea:#3a1f1f;--bg-fbf2da:#3a3216;--bg-fff:#17251d;--bg-fffcf2:#2a2614;--bg-ffffff:#17251d;--txt-0b3d2e:#e8f3ec;--txt-1c2a22:#dfece4;--txt-223028:#dfece4;--txt-2e7d5b:#5fbf93;--txt-3a5044:#c5d8cc;--txt-5b7768:#9bb5a6;--txt-6b5817:#e8d27a;--txt-7a5f0a:#f0d77a;--txt-7c9186:#8aa396;--txt-8a2f2f:#ff9a9a;--txt-9a7b0a:#e0c24d;--txt-a9bdb2:#6b8577;color-scheme:dark}
@media print{:root[data-theme="dark"]{--bd-0b3d2e:#0b3d2e;--bd-cfe0d6:#cfe0d6;--bd-dfe8e2:#dfe8e2;--bd-e4f2e9:#e4f2e9;--bd-e7eee9:#e7eee9;--bd-e8dcae:#e8dcae;--bd-eef3f0:#eef3f0;--bd-f2f7f4:#f2f7f4;--bd-fff:#fff;--bg-cfe0d6:#cfe0d6;--bg-e4f2e9:#e4f2e9;--bg-eef3f0:#eef3f0;--bg-efefef:#efefef;--bg-f2f7f4:#f2f7f4;--bg-f7f9f6:#f7f9f6;--bg-fafcf9:#fafcf9;--bg-fbeaea:#fbeaea;--bg-fbf2da:#fbf2da;--bg-fff:#fff;--bg-fffcf2:#fffcf2;--bg-ffffff:#ffffff;--txt-0b3d2e:#0b3d2e;--txt-1c2a22:#1c2a22;--txt-223028:#223028;--txt-2e7d5b:#2e7d5b;--txt-3a5044:#3a5044;--txt-5b7768:#5b7768;--txt-6b5817:#6b5817;--txt-7a5f0a:#7a5f0a;--txt-7c9186:#7c9186;--txt-8a2f2f:#8a2f2f;--txt-9a7b0a:#9a7b0a;--txt-a9bdb2:#a9bdb2}}
body { background: var(--bg-f7f9f6); transition: background .2s; font-size: 16px; line-height: 1.55; }
body, input, select, textarea, button { font-family: 'Inter', Arial, sans-serif; }
input, select, textarea { line-height: 1.45; }
button { line-height: 1.35; }

:root[data-theme="dark"] input, :root[data-theme="dark"] select, :root[data-theme="dark"] textarea { background: #1d2e25 !important; border-color: #355044 !important; color: #e8f3ec !important; }
:root[data-theme="dark"] input::placeholder, :root[data-theme="dark"] textarea::placeholder { color: #6f897a; }
:root[data-theme="dark"] option { background: #1d2e25; color: #e8f3ec; }
:root[data-theme="dark"] input[type="range"] { background: transparent !important; }
:root[data-theme="dark"] ::selection { background: #2e7d5b; color: #fff; }
:root[data-theme="dark"] input:focus, :root[data-theme="dark"] select:focus, :root[data-theme="dark"] textarea:focus { border-color: #5fbf93 !important; box-shadow: 0 0 0 3px rgba(95,191,147,.25); }
:root[data-theme="dark"] img { filter: brightness(.92); }
:root[data-theme="dark"] [style*="box-shadow"] { box-shadow: 0 6px 24px rgba(0,0,0,.45) !important; }
@media print { :root[data-theme="dark"] input, :root[data-theme="dark"] select, :root[data-theme="dark"] textarea { background: #fff !important; color: #000 !important; border-color: #ccc !important; } :root[data-theme="dark"] img { filter: none; } }
* { box-sizing: border-box; }
body, input, select, textarea, button { font-family: 'Inter', sans-serif; }
input:focus, select:focus, textarea:focus { outline: none; border-color: #2E7D5B !important; box-shadow: 0 0 0 3px rgba(46,125,91,.22); }
button { cursor: pointer; font-family: inherit; }
table { border-collapse: collapse; width: 100%; }
th { text-align: left; font-size: 13.5px; text-transform: uppercase; letter-spacing: 0.4px; color: var(--txt-5b7768); padding: 10px 12px; border-bottom: 2px solid var(--bd-e4f2e9); }
td { padding: 12px 14px; font-size: 15.5px; border-bottom: 1px solid var(--bd-eef3f0); color: var(--txt-223028); }
tr:hover td { background: var(--bg-fafcf9); }
::-webkit-scrollbar { width: 8px; height: 8px; }
::-webkit-scrollbar-thumb { background: var(--bg-cfe0d6); border-radius: 8px; }
.crystalSpin { animation: spin 1.6s linear infinite; }
@keyframes spin { to { transform: rotate(360deg); } }
.spinIcon { animation: spin 0.8s linear infinite; }
.onlyMobile { display: none; }
@media (max-width: 880px) {
  .onlyMobile { display: inline-flex !important; }
}
@media (max-width: 880px) {
  aside[style] { position: fixed !important; left: -280px; top: 0; bottom: 0; transition: left .2s; }
  aside.mobileNavOpen { left: 0 !important; }
}
@media (max-width: 640px) { main[style] { padding: 20px 16px 48px !important; } }

/* GeoAcervo — nova identidade visual do painel */
.geoDashboard { max-width: 1680px; margin: 0 auto; }
.geoBrandText { display:flex; flex-direction:column; gap:2px; min-width:0; }
.geoBrandText strong { color:#fff; font-family:Fraunces,serif; font-size:21px; line-height:1.1; }
.geoBrandText small { color:#cfe0d6; font-size:10.5px; white-space:nowrap; }
.geoWelcome { min-height:174px; position:relative; overflow:hidden; display:flex; align-items:center; border-radius:17px; padding:28px 32px; margin:0 0 20px; background:linear-gradient(105deg,#e8f1ee 0%,#dcebe7 48%,#a9c9b8 100%); border:1px solid #d7e5dd; }
.geoWelcome:after { content:""; position:absolute; inset:0; background:linear-gradient(90deg,rgba(238,246,242,.98) 0%,rgba(231,242,237,.92) 46%,rgba(231,242,237,.08) 100%); pointer-events:none; }
.geoWelcomeText { position:relative; z-index:2; }
.geoEyebrow { color:#39785b; font-size:11px; font-weight:800; letter-spacing:1.6px; }
.geoWelcome h1 { margin:7px 0 5px; color:#103d2c; font-family:Fraunces,serif; font-size:clamp(25px,3vw,37px); line-height:1.15; }
.geoWelcome p { margin:0; color:#304b3d; font-size:17px; }
.geoWelcomeLine { display:block; width:58px; height:5px; border-radius:5px; background:#438c65; margin-top:18px; }
.geoWelcomeArt { position:absolute; z-index:1; right:7%; bottom:-32px; color:rgba(27,86,59,.32); transform:scaleX(1.55); }
.geoStatsGrid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:13px; margin-bottom:17px; }
.geoStatCard { min-width:0; min-height:104px; display:flex; align-items:center; gap:14px; position:relative; overflow:hidden; background:var(--bg-fff); border:1px solid var(--bd-e7eee9); border-radius:15px; padding:17px; box-shadow:0 2px 8px rgba(20,55,38,.035); }
.geoStatIcon { width:50px; height:50px; border-radius:13px; display:flex; align-items:center; justify-content:center; flex-shrink:0; }
.geoStatIcon.mineral { background:#dff2e6; color:#17643e; }.geoStatIcon.rock { background:#f1e8df; color:#79502b; }.geoStatIcon.recent { background:#e3eff9; color:#246b9a; }.geoStatIcon.total { background:#eee8f8; color:#6744a0; }
.geoStatCard span { display:block; color:var(--txt-5b7768); font-size:13px; font-weight:600; white-space:normal; }
.geoStatCard strong { display:block; margin-top:2px; color:var(--txt-1c2a22); font-size:29px; line-height:1.1; font-weight:800; }
.geoStatWatermark { position:absolute; right:13px; bottom:12px; opacity:.07; color:var(--txt-0b3d2e); }
.geoDashboardColumns { display:grid; grid-template-columns:minmax(310px,.88fr) minmax(420px,1.22fr); gap:15px; align-items:start; }
.geoDashboardLeft,.geoDashboardRight { display:flex; flex-direction:column; gap:15px; min-width:0; }
.geoPanel { min-width:0; background:var(--bg-fff); border:1px solid var(--bd-e7eee9); border-radius:16px; padding:18px; box-shadow:0 2px 8px rgba(20,55,38,.025); }
.geoPanelTitle { display:flex; align-items:center; gap:10px; color:var(--txt-0b3d2e); margin-bottom:15px; min-width:0; }
.geoPanelTitle h2 { font-family:Fraunces,serif; font-size:18px; margin:0; font-weight:700; color:var(--txt-1c2a22); }
.geoPanelTitle > svg { flex-shrink:0; }
.geoZap { display:inline-flex; align-items:center; justify-content:center; width:23px; height:23px; color:#17643e; font-size:28px; font-weight:900; line-height:1; }
.geoQuickActions { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:10px; }
.geoQuickButton { min-height:78px; border:0; border-radius:12px; padding:12px; display:flex; align-items:center; gap:10px; text-align:left; color:#fff; font-weight:700; font-size:14px; line-height:1.3; transition:transform .15s,filter .15s; }
.geoQuickButton:hover { transform:translateY(-2px); filter:brightness(1.06); }
.geoQuickButton.dark { background:linear-gradient(135deg,#0b3d2e,#15573d); }.geoQuickButton.green { background:linear-gradient(135deg,#438c65,#286b49); }.geoQuickButton.soft { background:var(--bg-e4f2e9); color:var(--txt-0b3d2e); }
.geoQuickButton span { flex:1; }.geoQuickButton > svg:last-child { flex-shrink:0; opacity:.8; }
.geoQuickHint { grid-column:1/-1; color:var(--txt-5b7768); font-size:14px; margin:4px 0; }
.geoPanelNote { margin-left:auto; font-size:11px; color:var(--txt-7c9186); }
.geoGalleryGrid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:8px; }
.geoGalleryItem { min-width:0; overflow:hidden; border:0; padding:0; background:var(--bg-f2f7f4); border-radius:10px; text-align:left; color:var(--txt-1c2a22); }
.geoGalleryItem img { display:block; width:100%; height:96px; object-fit:cover; }.geoGalleryItem span { display:block; padding:7px 8px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-size:12px; font-weight:700; }
.geoGalleryEmpty { min-height:88px; display:flex; align-items:center; justify-content:center; gap:10px; color:var(--txt-7c9186); font-size:13px; background:var(--bg-f2f7f4); border-radius:10px; padding:15px; }
.geoRecentPanel { padding-bottom:8px; }
.geoSeeAll { margin-left:auto; display:inline-flex; align-items:center; gap:4px; border:0; background:transparent; color:#28734d; font-size:12px; font-weight:700; white-space:nowrap; }
.geoRecentList { display:flex; flex-direction:column; }
.geoRecentRow { display:grid; grid-template-columns:48px minmax(100px,1.4fr) minmax(58px,.7fr) minmax(75px,.7fr) 17px; gap:10px; align-items:center; width:100%; border:0; border-top:1px solid var(--bd-eef3f0); padding:9px 2px; background:transparent; text-align:left; color:var(--txt-1c2a22); }
.geoRecentRow:hover { background:var(--bg-f2f7f4); }
.geoRecentThumb { width:44px; height:40px; border-radius:7px; background:var(--bg-f2f7f4); color:var(--txt-2e7d5b); display:flex; align-items:center; justify-content:center; overflow:hidden; }
.geoRecentThumb img { width:100%; height:100%; object-fit:cover; }
.geoRecentName { min-width:0; font-size:13.5px; font-weight:700; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.geoRecentName small { display:block; color:var(--txt-7c9186); font-family:'JetBrains Mono',monospace; font-size:10.5px; font-weight:500; margin-top:3px; }
.geoRecentType,.geoRecentDate { color:var(--txt-5b7768); font-size:12px; white-space:nowrap; }
.geoRecentArrow { color:var(--txt-7c9186); }
.geoKnowledgeBanner { min-height:108px; position:relative; overflow:hidden; display:flex; align-items:center; gap:15px; padding:18px 22px; border-radius:15px; background:linear-gradient(105deg,#dcebe4,#c3dbcf); color:#143f2d; }
.geoKnowledgeIcon { width:48px; height:48px; border-radius:13px; display:flex; align-items:center; justify-content:center; background:#14543a; color:white; flex-shrink:0; }
.geoKnowledgeBanner h2 { margin:0 0 4px; font-family:Fraunces,serif; font-size:17px; }.geoKnowledgeBanner p { max-width:410px; margin:0; font-size:12.5px; line-height:1.5; color:#365647; }
.geoKnowledgeArt { margin-left:auto; flex-shrink:0; color:rgba(31,91,61,.22); }
@media(max-width:1180px) { .geoStatsGrid { grid-template-columns:repeat(2,minmax(0,1fr)); }.geoDashboardColumns { grid-template-columns:1fr; }.geoDashboardLeft { display:grid; grid-template-columns:1fr 1fr; align-items:stretch; } }
@media(max-width:760px) { .geoWelcome { min-height:150px; padding:22px 20px; }.geoWelcome p { font-size:14px; max-width:75%; }.geoWelcomeArt { right:-4%; opacity:.65; }.geoStatsGrid { gap:9px; }.geoStatCard { min-height:90px; padding:12px; gap:9px; }.geoStatIcon { width:40px; height:40px; }.geoStatCard span { font-size:11px; }.geoStatCard strong { font-size:24px; }.geoDashboardLeft { display:flex; }.geoPanel { padding:15px; }.geoRecentRow { grid-template-columns:42px minmax(90px,1fr) 56px 17px; }.geoRecentDate { display:none; }.geoRecentType { font-size:11px; }.geoGalleryItem img { height:78px; } }
@media(max-width:430px) { .geoWelcome h1 { font-size:23px; }.geoWelcome p { max-width:90%; font-size:13px; }.geoStatsGrid { grid-template-columns:repeat(2,minmax(0,1fr)); }.geoStatCard { align-items:flex-start; flex-direction:column; gap:6px; }.geoStatIcon { width:34px; height:34px; }.geoStatIcon svg { width:18px; }.geoStatCard strong { font-size:22px; }.geoQuickActions { grid-template-columns:1fr 1fr; }.geoQuickButton { padding:10px; font-size:12px; gap:7px; }.geoPanelTitle h2 { font-size:16px; }.geoGalleryGrid { grid-template-columns:repeat(2,minmax(0,1fr)); }.geoGalleryItem img { height:100px; }.geoRecentRow { gap:7px; grid-template-columns:38px minmax(75px,1fr) 52px 14px; }.geoRecentThumb { width:36px; height:36px; }.geoRecentName { font-size:12px; }.geoRecentType { font-size:10px; }.geoKnowledgeArt { display:none; } }

/* Navegação mobile simples e direta */
.mobileBrand { display:none; }
.mobileBottomNav { display:none; }
@media(max-width:640px) {
  .mobileBrand { display:flex; align-items:center; gap:9px; margin-right:auto; min-width:0; }
  .mobileBrand strong { display:block; color:var(--txt-0b3d2e); font-size:16px; line-height:1.15; }
  .mobileBrand small { display:block; color:var(--txt-5b7768); font-size:11px; }
  .topbar[style] { height:auto !important; min-height:62px; padding:10px 14px !important; display:flex; flex-wrap:wrap; gap:10px !important; }
  .topbar .userChip, .topbar > button[title="Atualizar dados agora"], .topbar > .onlyMobile { display:none !important; }
  .topbar > div[style*="position: relative"] { order:2; flex:1 0 100%; max-width:none !important; }
  .topbar > div[style*="position: relative"] input { font-size:14px !important; }
  aside[style] { width: min(290px, 86vw) !important; left: calc(-1 * min(300px, 90vw)) !important; z-index:90 !important; box-shadow:12px 0 35px rgba(0,0,0,.18); }
  aside.mobileNavOpen { left:0 !important; }
  .mobileOverlay { z-index:80 !important; }
  .main { width:100%; min-width:0; }
  main[style] { padding:16px 14px 100px !important; }
  .mobileBottomNav { position:fixed; z-index:70; left:0; right:0; bottom:0; display:grid; grid-template-columns:repeat(4,1fr); gap:2px; padding:7px 8px calc(7px + env(safe-area-inset-bottom)); background:var(--bg-fff); border-top:1px solid var(--bd-e7eee9); box-shadow:0 -5px 18px rgba(15,45,30,.08); }
  .mobileBottomNav button { min-width:0; min-height:54px; border:0; border-radius:11px; background:transparent; color:var(--txt-5b7768); display:flex; flex-direction:column; align-items:center; justify-content:center; gap:3px; font-size:11px; font-weight:600; }
  .mobileBottomNav button.active { color:#fff; background:#0b3d2e; }
  .geoWelcome { min-height:0; padding:17px 16px; margin-bottom:13px; border-radius:13px; }
  .geoWelcome:after { background:linear-gradient(90deg,rgba(238,246,242,.98),rgba(231,242,237,.88)); }
  .geoWelcome h1 { font-size:21px; margin:5px 0; }
  .geoWelcome p { font-size:13px; max-width:100%; }
  .geoWelcomeLine, .geoWelcomeArt, .geoEyebrow { display:none; }
  .geoStatsGrid { grid-template-columns:repeat(2,minmax(0,1fr)); gap:9px; margin-bottom:13px; }
  .geoStatCard { min-height:76px; padding:11px; display:flex; flex-direction:row; align-items:center; gap:9px; border-radius:12px; }
  .geoStatCard:nth-child(n+3) { display:none; }
  .geoStatIcon { width:36px; height:36px; border-radius:10px; }
  .geoStatIcon svg { width:18px; height:18px; }
  .geoStatCard span { font-size:10.5px; line-height:1.25; }
  .geoStatCard strong { font-size:22px; }
  .geoStatWatermark { display:none; }
  .geoDashboardColumns { display:flex; flex-direction:column; gap:12px; }
  .geoDashboardLeft, .geoDashboardRight { display:contents; }
  .geoQuickPanel { order:0; }
  .geoRecentPanel { order:1; }
  .geoGalleryPanel, .geoKnowledgeBanner { display:none !important; }
  .geoPanel { padding:13px; border-radius:12px; margin:0; }
  .geoPanelTitle { margin-bottom:10px; }
  .geoPanelTitle h2 { font-family:Inter,sans-serif; font-size:15px; }
  .geoPanelNote { display:none; }
  .geoQuickActions { grid-template-columns:repeat(2,minmax(0,1fr)); gap:8px; }
  .geoQuickButton { min-height:60px; padding:10px; border-radius:10px; font-size:12px; gap:7px; }
  .geoQuickButton > svg:first-child { width:20px; height:20px; }
  .geoQuickButton.soft { display:none; }
  .geoRecentRow { grid-template-columns:42px minmax(0,1fr) auto 14px; gap:8px; padding:10px 0; }
  .geoRecentThumb { width:40px; height:40px; }
  .geoRecentName { font-size:13px; }
  .geoRecentName small { font-size:10px; }
  .geoRecentType { font-size:10px; }
  .geoRecentDate { display:none; }
  .geoSeeAll { font-size:11px; }
  .card { padding:14px !important; border-radius:12px !important; }
  .filterGrid { grid-template-columns:1fr !important; }
  .modalOverlay { padding:0 !important; }
  .modal { margin:0 !important; min-height:100dvh; border-radius:0 !important; }
  .modalHead { padding:14px 16px !important; }
  .modalBody { padding:8px 14px 0 !important; }
  .modalFooter { padding:12px 14px !important; margin:20px -14px 0 !important; }
  input, select, textarea { font-size:16px !important; }
  table { min-width:620px; }
}
@media(max-width:380px) {
  .geoStatCard { padding:9px; gap:7px; }
  .geoStatCard span { font-size:10px; }
  .geoStatCard strong { font-size:20px; }
  .geoQuickButton { font-size:11px; }
}

/* Formulários de minerais e rochas: experiência mobile acessível */
@media (max-width: 640px) {
  .geoFormModal {
    width: 100% !important;
    max-width: none !important;
    height: 100dvh !important;
    min-height: 100dvh !important;
    max-height: 100dvh !important;
    margin: 0 !important;
    display: flex !important;
    flex-direction: column !important;
    overflow: hidden !important;
    border-radius: 0 !important;
  }
  .geoFormModal .modalHead {
    flex: 0 0 auto;
    min-height: 60px;
    padding: 12px 15px !important;
    border-radius: 0 !important;
    box-shadow: 0 2px 8px rgba(10,45,29,.06);
  }
  .geoFormModal .modalHead h3 { font-size: 17px !important; line-height: 1.3 !important; }
  .geoFormModal .modalHead button { width: 40px !important; height: 40px !important; flex: 0 0 40px; }
  .geoFormModalBody {
    flex: 1 1 auto;
    min-height: 0;
    overflow-y: auto !important;
    overflow-x: hidden !important;
    -webkit-overflow-scrolling: touch;
    overscroll-behavior: contain;
    padding: 4px 13px 18px !important;
  }
  .geoFormModalBody > div { min-width: 0; }
  .geoFormModalBody .sectionTitle,
  .geoFormModalBody [style*="font-size: 13.5px"][style*="text-transform: uppercase"] {
    display: flex;
    align-items: center;
    gap: 9px;
    margin: 22px 0 10px !important;
    padding: 11px 12px;
    border-left: 4px solid #2e7d5b;
    border-radius: 8px;
    background: var(--bg-e4f2e9);
    color: var(--txt-0b3d2e) !important;
    font-size: 13px !important;
    line-height: 1.4;
    letter-spacing: .2px !important;
  }
  .geoFormModalBody .formGrid2,
  .geoFormModalBody .formGrid3,
  .geoFormModalBody .formGrid4 {
    display: grid !important;
    grid-template-columns: minmax(0, 1fr) !important;
    gap: 13px !important;
    padding: 13px !important;
    margin: 0 0 13px !important;
    border-radius: 11px !important;
    background: var(--bg-fafcf9) !important;
  }
  .geoFormModalBody .formGrid2 > *,
  .geoFormModalBody .formGrid3 > *,
  .geoFormModalBody .formGrid4 > * { min-width: 0; grid-column: 1 / -1 !important; }
  .geoFormModalBody label {
    display: block;
    margin-bottom: 6px !important;
    font-size: 13px !important;
    line-height: 1.4;
    font-weight: 700 !important;
    text-transform: none !important;
    letter-spacing: 0 !important;
    color: var(--txt-1c2a22) !important;
  }
  .geoFormModalBody input:not([type="checkbox"]):not([type="radio"]),
  .geoFormModalBody select,
  .geoFormModalBody textarea {
    display: block;
    width: 100% !important;
    max-width: 100% !important;
    min-width: 0 !important;
    min-height: 48px;
    box-sizing: border-box !important;
    padding: 11px 12px !important;
    border-radius: 9px !important;
    font-size: 16px !important;
    line-height: 1.4 !important;
    background-color: var(--bg-fff) !important;
    color: var(--txt-1c2a22) !important;
    -webkit-appearance: auto;
  }
  .geoFormModalBody textarea { min-height: 104px !important; resize: vertical; }
  .geoFormModalBody .mohsRow { display: grid !important; grid-template-columns: repeat(5,minmax(0,1fr)); gap: 7px !important; }
  .geoFormModalBody .mohsFacet { height: 38px !important; clip-path: none !important; border-radius: 8px !important; }
  .geoFormModalBody .photoDrop { min-height: 54px; padding: 13px !important; border-radius: 10px; flex-wrap: wrap; }
  .geoFormModalBody .modalFooter {
    position: sticky !important;
    bottom: -18px !important;
    display: grid !important;
    grid-template-columns: minmax(0,.8fr) minmax(0,1.2fr) !important;
    gap: 9px !important;
    margin: 20px -13px -18px !important;
    padding: 11px 13px calc(11px + env(safe-area-inset-bottom)) !important;
    background: var(--bg-fff) !important;
    box-shadow: 0 -5px 16px rgba(10,45,29,.08);
    border-radius: 0 !important;
  }
  .geoFormModalBody .modalFooter button { min-height: 47px; width: 100% !important; justify-content: center; padding: 10px 9px !important; font-size: 14px !important; }
  .geoFormModalBody .tipoHint { align-items: flex-start; line-height: 1.5; padding: 12px; }
  .geoFormModalBody .qapWrap { padding: 11px !important; overflow: hidden; }
  .geoFormModalBody .qapLegendGrid { grid-template-columns: minmax(0,1fr) !important; }
  .geoFormModalBody .thumbWrap { width: 72px; height: 72px; }
  .geoFormModalBody .thumb { width: 72px; height: 72px; }
}
/* Refinamento mobile dos formulários GeoAcervo — identidade verde original */
.geoMobileFormIntro { display:none; }
@media (max-width: 640px) {
  .geoMobileFormIntro { display:flex; align-items:center; gap:12px; margin:8px 0 18px; padding:13px; border:1px solid var(--bd-cfe0d6); border-radius:12px; background:linear-gradient(135deg,var(--bg-e4f2e9),var(--bg-fafcf9)); }
  .geoMobileFormIntroIcon { display:flex; align-items:center; justify-content:center; flex:0 0 42px; width:42px; height:42px; border-radius:11px; color:#fff; background:linear-gradient(135deg,#2e7d5b,#0b3d2e); }
  .geoMobileFormIntro strong { display:block; color:var(--txt-0b3d2e); font-size:15px; line-height:1.35; }
  .geoMobileFormIntro small { display:block; margin-top:3px; color:var(--txt-5b7768); font-size:12px; line-height:1.45; }
  .geoFormModalBody .sectionTitle { position:relative; margin-top:24px !important; padding:12px 13px !important; border:1px solid var(--bd-cfe0d6); border-left:4px solid #2e7d5b; background:var(--bg-e4f2e9); color:var(--txt-0b3d2e) !important; border-radius:10px; font-size:14px !important; }
  .geoFormModalBody .geoFormField { min-width:0; }
  .geoFormModalBody .geoFormFieldLabel { display:block; margin-bottom:7px !important; color:var(--txt-1c2a22) !important; font-size:13px !important; font-weight:700 !important; line-height:1.4 !important; text-transform:none !important; letter-spacing:0 !important; }
  .geoFormModalBody .formGrid2, .geoFormModalBody .formGrid3, .geoFormModalBody .formGrid4 { border:1px solid var(--bd-e7eee9) !important; box-shadow:0 2px 7px rgba(11,61,46,.035); }
  .geoFormModalBody input:focus, .geoFormModalBody select:focus, .geoFormModalBody textarea:focus { outline:3px solid rgba(46,125,91,.18) !important; border-color:#2e7d5b !important; }
  .geoFormModalBody .photoDrop { border:1.5px dashed #8db6a0 !important; background:var(--bg-e4f2e9) !important; color:var(--txt-0b3d2e) !important; }
  .geoFormModalBody .modalFooter { border-top:1px solid var(--bd-e7eee9); }
  .geoFormModalBody .modalFooter button:last-child { background:linear-gradient(135deg,#2e7d5b,#0b3d2e) !important; box-shadow:0 3px 9px rgba(11,61,46,.16); }
}

@media (max-width: 380px) {
  .geoFormModalBody { padding-left: 10px !important; padding-right: 10px !important; }
  .geoFormModalBody .formGrid2,
  .geoFormModalBody .formGrid3,
  .geoFormModalBody .formGrid4 { padding: 10px !important; }
  .geoFormModalBody .modalFooter { margin-left: -10px !important; margin-right: -10px !important; }
}

@media print {
  body * { visibility: hidden; }
  #fichaPrint, #fichaPrint * { visibility: visible; }
  #fichaPrint { position: absolute; top: 0; left: 0; width: 100%; box-shadow: none !important; }
  .noPrint { display: none !important; }
}

/* Navegação por etapas dos formulários */
.geoWizardProgress { margin: 14px 0 18px; padding: 13px 14px; border: 1px solid var(--bd-dfe8e2); border-radius: 12px; background: var(--bg-fafcf9); }
.geoWizardTop { display:flex; align-items:center; justify-content:space-between; gap:10px; margin-bottom:10px; color:var(--txt-5b7768); font-size:12px; }
.geoWizardTop strong { color:var(--txt-0b3d2e); font-size:13px; }
.geoWizardTrack { height:5px; background:var(--bg-e4f2e9); border-radius:99px; overflow:hidden; }
.geoWizardTrack span { display:block; height:100%; border-radius:99px; background:#2e7d5b; transition:width .2s ease; }
.geoWizardSteps { display:flex; justify-content:space-between; margin-top:10px; }
.geoWizardSteps span { width:23px; height:23px; display:grid; place-items:center; border-radius:50%; background:var(--bg-e4f2e9); color:var(--txt-5b7768); font-size:11px; font-weight:700; }
.geoWizardSteps span.active,.geoWizardSteps span.done { background:#2e7d5b; color:#fff; }
.geoWizardPanel[hidden] { display:none !important; }
.geoWizardPanel { width:100%; min-width:0; }
.geoWizardActions { display:flex; align-items:center; justify-content:space-between; gap:10px; padding:13px 0 4px; margin-top:18px; border-top:1px solid var(--bd-dfe8e2); }
.geoWizardActions button { display:inline-flex; align-items:center; justify-content:center; gap:7px; min-height:44px; }
@media(max-width:640px) {
  /* O modal fica acima da navegação inferior para que Salvar seja clicável. */
  .geoFormModal { max-height: calc(100dvh - 12px) !important; height: calc(100dvh - 12px) !important; margin: 6px 0 !important; display:flex !important; flex-direction:column !important; min-width:0 !important; width:100% !important; }
  .geoFormModalBody { padding: 4px 13px 24px !important; }
  .geoWizardProgress { margin:10px 0 14px; padding:11px 12px; }
  .geoWizardActions { position:sticky; bottom:0; z-index:8; margin:14px -13px 0; padding:12px 13px calc(14px + env(safe-area-inset-bottom)); background:var(--bg-fff); box-shadow:0 -5px 16px rgba(11,61,46,.10); }
  .geoWizardActions button { flex:1; min-width:0; }
  .geoFormModalBody .geoFormField,
  .geoFormModalBody .geoFormField input,
  .geoFormModalBody .geoFormField select,
  .geoFormModalBody .geoFormField textarea { box-sizing:border-box !important; width:100%; max-width:100% !important; min-width:0 !important; }
  .geoFormModalBody .formGrid2,
  .geoFormModalBody .formGrid3,
  .geoFormModalBody .formGrid4 { width:100%; min-width:0 !important; box-sizing:border-box !important; }
}

`;

const S = {
  app: { display: "flex", minHeight: "100vh", background: "var(--bg-f7f9f6)", fontFamily: "Inter" },
  bootWrap: { minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "var(--bg-f7f9f6)" },
  bootCard: { textAlign: "center" },
  crystal: { width: 46, height: 46, background: "linear-gradient(135deg,#2E7D5B,#0B3D2E)", clipPath: "polygon(50% 0%, 100% 38%, 82% 100%, 18% 100%, 0% 38%)", margin: "0 auto" },

  loginWrap: { minHeight: "100vh", position: "relative", display: "flex", alignItems: "center", justifyContent: "center", background: "#0B3D2E", overflow: "hidden" },
  loginBg: { position: "absolute", inset: 0, background: "radial-gradient(circle at 20% 20%, #1B5E42 0%, #0B3D2E 55%, #072518 100%)" },
  loginCard: { position: "relative", background: "var(--bg-ffffff)", borderRadius: 18, padding: "36px 38px", width: 400, maxWidth: "90vw", boxShadow: "0 30px 80px rgba(0,0,0,0.35)" },
  loginBrand: { display: "flex", alignItems: "center", gap: 12 },
  hexIcon: { width: 46, height: 46, background: "linear-gradient(135deg,#2E7D5B,#0B3D2E)", clipPath: "polygon(50% 0%, 100% 25%, 100% 75%, 50% 100%, 0% 75%, 0% 25%)", display: "flex", alignItems: "center", justifyContent: "center" },
  hexIconSm: { width: 34, height: 34, background: "var(--bg-f7f9f6)", clipPath: "polygon(50% 0%, 100% 25%, 100% 75%, 50% 100%, 0% 75%, 0% 25%)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 },
  label: { display: "block", fontSize: 14.5, fontWeight: 600, color: "var(--txt-0b3d2e)", marginBottom: 6, marginTop: 16 },
  labelSm: { display: "block", fontSize: 13.5, fontWeight: 600, color: "var(--txt-5b7768)", marginBottom: 6, textTransform: "uppercase", letterSpacing: 0.3 },
  input: { width: "100%", height: 44, padding: "0 13px", borderRadius: 10, transition: "border-color .15s", border: "1.5px solid var(--bd-dfe8e2)", fontSize: 16, background: "var(--bg-fff)", color: "var(--txt-1c2a22)" },
  inputSm: { width: "100%", padding: "8px 10px", borderRadius: 8, border: "1.5px solid var(--bd-dfe8e2)", fontSize: 15, background: "var(--bg-fff)" },
  loginErr: { display: "flex", alignItems: "center", gap: 6, color: "var(--txt-8a2f2f)", fontSize: 14.5, marginTop: 10, background: "var(--bg-fbeaea)", padding: "8px 10px", borderRadius: 8 },
  btnPrimary: { display: "flex", alignItems: "center", justifyContent: "center", gap: 8, width: "100%", marginTop: 22, background: "#0B3D2E", color: "#fff", border: "none", padding: "12px 16px", borderRadius: 10, fontWeight: 600, fontSize: 15.5 },
  btnPrimarySm: { display: "flex", alignItems: "center", gap: 7, background: "#0B3D2E", color: "#fff", border: "none", padding: "9px 15px", borderRadius: 9, fontWeight: 600, fontSize: 15 },
  btnGhost: { display: "flex", alignItems: "center", justifyContent: "center", gap: 8, background: "var(--bg-fff)", color: "var(--txt-0b3d2e)", border: "1.5px solid var(--bd-dfe8e2)", padding: "10px 16px", borderRadius: 9, fontWeight: 600, fontSize: 15.5 },
  btnGhostSm: { display: "flex", alignItems: "center", gap: 6, background: "var(--bg-fff)", color: "var(--txt-0b3d2e)", border: "1.5px solid var(--bd-dfe8e2)", padding: "8px 13px", borderRadius: 8, fontWeight: 600, fontSize: 14.5 },
  loginHint: { marginTop: 22, fontSize: 13.5, color: "var(--txt-5b7768)", background: "var(--bg-f2f7f4)", padding: "10px 12px", borderRadius: 9, lineHeight: 1.6 },
  demoBtn: { display: "inline-flex", alignItems: "center", gap: 6, background: "var(--bg-fff)", border: "1px solid var(--bd-cfe0d6)", color: "var(--txt-0b3d2e)", padding: "6px 10px", borderRadius: 7, fontSize: 13.5, fontWeight: 600 },
  authTabs: { display: "flex", gap: 4, marginTop: 24, background: "var(--bg-f2f7f4)", borderRadius: 10, padding: 4 },
  authTab: { flex: 1, background: "transparent", border: "none", padding: "9px 0", borderRadius: 7, fontSize: 15, fontWeight: 600, color: "var(--txt-5b7768)" },
  authTabActive: { background: "var(--bg-fff)", color: "var(--txt-0b3d2e)", boxShadow: "0 2px 6px rgba(11,61,46,0.12)" },
  regOkBox: { textAlign: "center", padding: "24px 10px 8px" },
  forgotLink: { display: "block", width: "100%", textAlign: "center", background: "transparent", border: "none", color: "var(--txt-5b7768)", fontSize: 14.5, fontWeight: 600, marginTop: 12, textDecoration: "underline" },

  sidebar: { background: "#0B3D2E", display: "flex", flexDirection: "column", transition: "width .18s", position: "sticky", top: 0, height: "100vh", flexShrink: 0, zIndex: 20 },
  sidebarTop: { display: "flex", alignItems: "center", gap: 10, padding: "20px 18px", borderBottom: "1px solid rgba(255,255,255,0.08)" },
  nav: { display: "flex", flexDirection: "column", gap: 3, padding: "14px 10px", flex: 1, overflowY: "auto" },
  navItem: { display: "flex", alignItems: "center", gap: 11, background: "transparent", border: "none", color: "#cfe0d6", padding: "10px 12px", borderRadius: 9, fontSize: 15.5, fontWeight: 500, textAlign: "left" },
  navItemActive: { background: "rgba(255,255,255,0.12)", color: "#fff" },
  navBadge: { background: "#C9A227", color: "#fff", fontSize: 12.5, fontWeight: 700, borderRadius: 999, padding: "1px 7px", minWidth: 18, textAlign: "center" },
  sidebarBottom: { padding: 12, borderTop: "1px solid rgba(255,255,255,0.08)", display: "flex", flexDirection: "column", gap: 8 },
  sidebarCollapse: { background: "rgba(255,255,255,0.08)", border: "none", color: "#cfe0d6", borderRadius: 8, width: 34, height: 34, display: "flex", alignItems: "center", justifyContent: "center" },
  logoutBtn: { flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 7, background: "rgba(255,255,255,0.08)", border: "none", color: "#f5d5d5", borderRadius: 8, padding: "8px 10px", fontSize: 15, fontWeight: 600 },
  overlay: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", zIndex: 15 },

  main: { flex: 1, minWidth: 0, display: "flex", flexDirection: "column" },
  topbar: { display: "flex", alignItems: "center", gap: 16, padding: "14px 26px", background: "var(--bg-fff)", borderBottom: "1px solid var(--bd-e7eee9)", position: "sticky", top: 0, zIndex: 10 },
  menuBtn: { background: "var(--bg-f2f7f4)", border: "none", borderRadius: 8, width: 36, height: 36, alignItems: "center", justifyContent: "center", color: "var(--txt-0b3d2e)" },
  refreshBtn: { background: "var(--bg-f2f7f4)", border: "none", borderRadius: 8, width: 36, height: 36, display: "flex", alignItems: "center", justifyContent: "center", color: "var(--txt-0b3d2e)", flexShrink: 0 },
  searchBox: { flex: 1, position: "relative", display: "flex", alignItems: "center", gap: 8, background: "var(--bg-f2f7f4)", borderRadius: 10, padding: "9px 14px", maxWidth: 640 },
  searchInput: { flex: 1, border: "none", background: "transparent", fontSize: 16, outline: "none" },
  searchResults: { position: "absolute", top: "110%", left: 0, right: 0, background: "var(--bg-fff)", borderRadius: 10, boxShadow: "0 14px 40px rgba(11,61,46,0.18)", border: "1px solid var(--bd-e7eee9)", maxHeight: 320, overflowY: "auto", zIndex: 30 },
  searchResultRow: { display: "flex", alignItems: "center", gap: 10, padding: "10px 14px", borderBottom: "1px solid var(--bd-f2f7f4)", fontSize: 15 },
  searchEmpty: { padding: 16, fontSize: 15, color: "var(--txt-5b7768)" },
  userChip: { display: "flex", alignItems: "center", gap: 10 },
  userAvatar: { width: 34, height: 34, borderRadius: 9, background: "#0B3D2E", color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 16 },

  content: { padding: "30px 34px 64px", flex: 1 },
  pageHeader: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 22, flexWrap: "wrap", gap: 12 },
  pageTitle: { fontFamily: "Fraunces", fontSize: 27, color: "var(--txt-0b3d2e)", margin: 0, fontWeight: 600 },
  pageSubtitle: { color: "var(--txt-5b7768)", fontSize: 15.5, marginTop: 4 },

  statsGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px,1fr))", gap: 14, marginBottom: 20 },
  statCard: { background: "var(--bg-fff)", border: "1px solid var(--bd-e7eee9)", borderRadius: 14, padding: "16px 18px", display: "flex", alignItems: "center", gap: 13 },
  statIcon: { width: 44, height: 44, borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 },
  statValue: { fontFamily: "Fraunces", fontSize: 25, fontWeight: 700, color: "var(--txt-0b3d2e)", lineHeight: 1 },
  statLabel: { fontSize: 14, color: "var(--txt-5b7768)", marginTop: 4 },
  statSub: { fontSize: 12.5, color: "var(--txt-a9bdb2)" },

  quickGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(210px,1fr))", gap: 12, marginBottom: 22 },
  quickCard: { display: "flex", alignItems: "center", gap: 12, background: "var(--bg-fff)", border: "1px solid var(--bd-e7eee9)", borderRadius: 12, padding: "14px 16px", fontSize: 15.5, fontWeight: 600, color: "var(--txt-1c2a22)" },
  quickIcon: { width: 34, height: 34, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 },

  card: { background: "var(--bg-fff)", border: "1px solid var(--bd-e7eee9)", borderRadius: 16, padding: 22, marginBottom: 18 },
  cardHeadRow: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 },
  cardTitle: { fontFamily: "Fraunces", fontSize: 17.5, color: "var(--txt-0b3d2e)", margin: "0 0 12px", fontWeight: 600 },
  tableWrap: { overflowX: "auto" },
  table: { fontSize: 15.5 },
  mono: { fontFamily: "JetBrains Mono", fontSize: 14, color: "var(--txt-0b3d2e)", fontWeight: 600 },
  badge: { display: "inline-flex", alignItems: "center", gap: 4, background: "var(--bg-f2f7f4)", color: "var(--txt-0b3d2e)", padding: "3px 9px", borderRadius: 999, fontSize: 13.5, fontWeight: 600 },
  iconBtn: { background: "var(--bg-f2f7f4)", border: "none", borderRadius: 7, width: 30, height: 30, display: "inline-flex", alignItems: "center", justifyContent: "center", color: "var(--txt-0b3d2e)" },
  emptyState: { textAlign: "center", padding: "40px 20px", color: "var(--txt-7c9186)", fontSize: 15.5, display: "flex", flexDirection: "column", alignItems: "center", gap: 10 },

  filterRow: { marginBottom: 16 },
  searchBoxSm: { display: "flex", alignItems: "center", gap: 8, background: "var(--bg-fff)", border: "1.5px solid var(--bd-e7eee9)", borderRadius: 10, padding: "9px 13px", maxWidth: 420 },
  searchInputSm: { flex: 1, border: "none", outline: "none", fontSize: 15 },
  filterGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px,1fr))", gap: 12 },

  cardGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(230px,1fr))", gap: 16 },
  sampleCard: { background: "var(--bg-fff)", border: "1px solid var(--bd-e7eee9)", borderRadius: 14, overflow: "hidden" },
  sampleImg: { position: "relative", height: 130, background: "var(--bg-eef3f0)" },
  sampleImgTag: { width: "100%", height: "100%", objectFit: "cover" },
  sampleImgPh: { width: "100%", height: "100%", display: "flex", alignItems: "center", justifyContent: "center" },
  codeChip: { position: "absolute", bottom: 8, left: 8, background: "var(--bg-fff)", border: "1.5px solid", fontFamily: "JetBrains Mono", fontSize: 13, fontWeight: 700, padding: "2px 8px", borderRadius: 999, color: "var(--txt-0b3d2e)" },

  reportGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(300px,1fr))", gap: 16, marginBottom: 4 },
  barRow: { display: "flex", alignItems: "center", gap: 12, marginBottom: 10, fontSize: 15 },
  barTrack: { flex: 1, height: 10, background: "var(--bg-eef3f0)", borderRadius: 6, overflow: "hidden" },
  barFill: { height: "100%", borderRadius: 6 },

  modalOverlay: { position: "fixed", inset: 0, background: "rgba(11,61,46,0.35)", display: "flex", alignItems: "flex-start", justifyContent: "center", padding: "0 16px", overflowY: "auto", zIndex: 110 },
  modal: { background: "var(--bg-fff)", borderRadius: 16, width: "100%", margin: "30px 0", boxShadow: "0 30px 80px rgba(0,0,0,0.3)" },
  modalHead: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "18px 24px", borderBottom: "1px solid var(--bd-eef3f0)", borderLeft: "5px solid #2E7D5B", borderRadius: "16px 16px 0 0", position: "sticky", top: 0, background: "var(--bg-fff)", zIndex: 5 },
  modalBody: { padding: "6px 24px 0" },
  modalFooter: { display: "flex", justifyContent: "flex-end", gap: 10, padding: "14px 24px", margin: "28px -24px 0", borderTop: "1px solid var(--bd-eef3f0)", position: "sticky", bottom: 0, background: "var(--bg-fff)", borderRadius: "0 0 16px 16px", zIndex: 5 },
  formGrid2: { display: "grid", alignItems: "end", gridTemplateColumns: "repeat(auto-fit,minmax(240px,1fr))", gap: 16, padding: 16, background: "var(--bg-f2f7f4)", border: "1px solid var(--bd-eef3f0)", borderRadius: 12 },
  formGrid3: { display: "grid", alignItems: "end", gridTemplateColumns: "repeat(auto-fit,minmax(190px,1fr))", gap: 16, padding: 16, background: "var(--bg-f2f7f4)", border: "1px solid var(--bd-eef3f0)", borderRadius: 12 },
  formGrid4: { display: "grid", alignItems: "end", gridTemplateColumns: "repeat(auto-fit,minmax(140px,1fr))", gap: 14, marginTop: 10, padding: 16, background: "var(--bg-f2f7f4)", border: "1px solid var(--bd-eef3f0)", borderRadius: 12 },
  sectionTitle: { display: "flex", alignItems: "center", gap: 10, fontSize: 15, fontWeight: 700, color: "var(--txt-0b3d2e)", letterSpacing: 0.2, margin: "30px 0 12px" },
  sectionIcon: { display: "inline-flex", alignItems: "center", justifyContent: "center", width: 28, height: 28, borderRadius: 9, background: "var(--bg-e4f2e9)", color: "var(--txt-2e7d5b)", flexShrink: 0 },
  tipoHint: { display: "flex", alignItems: "center", gap: 8, background: "var(--bg-fffcf2)", border: "1px solid var(--bd-e8dcae)", color: "var(--txt-7a5f0a)", fontSize: 14, padding: "10px 12px", borderRadius: 9, marginTop: 4 },
  qapWrap: { background: "var(--bg-fafcf9)", border: "1px solid var(--bd-e4f2e9)", borderRadius: 12, padding: "16px 14px", marginTop: 4 },
  qapLegendGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(210px, 1fr))", gap: "4px 14px", marginTop: 12, fontSize: 13.5, color: "var(--txt-3a5044)" },
  qapLegendItem: { lineHeight: 1.5 },

  photoDrop: { display: "flex", alignItems: "center", gap: 8, border: "1.5px dashed var(--bd-cfe0d6)", borderRadius: 10, padding: "12px 14px", fontSize: 14.5, color: "var(--txt-5b7768)", background: "var(--bg-fafcf9)" },
  thumbWrap: { position: "relative", width: 60, height: 60 },
  thumb: { width: 60, height: 60, objectFit: "cover", borderRadius: 8, border: "1px solid var(--bd-dfe8e2)" },
  thumbRemove: { position: "absolute", top: -6, right: -6, background: "#8a2f2f", color: "#fff", border: "2px solid var(--bd-fff)", borderRadius: "50%", width: 20, height: 20, display: "flex", alignItems: "center", justifyContent: "center" },

  mohsRow: { display: "flex", gap: 4, marginTop: 4 },
  mohsFacet: { flex: 1, height: 30, border: "none", borderRadius: 4, fontSize: 13.5, fontWeight: 700, clipPath: "polygon(15% 0,100% 0,85% 100%,0 100%)" },

  fichaHead: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 18, paddingBottom: 16, borderBottom: "2px solid var(--bd-e4f2e9)" },
  fichaSecTitulo: { fontSize: 13, fontWeight: 700, color: "var(--txt-2e7d5b)", textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 8 },
  fichaGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(150px,1fr))", gap: 10 },
  fichaLabel: { fontSize: 12.5, color: "var(--txt-a9bdb2)", textTransform: "uppercase", letterSpacing: 0.3 },
  fichaValor: { fontSize: 15.5, color: "var(--txt-1c2a22)", fontWeight: 600, marginTop: 2 },

  toast: { position: "fixed", bottom: 24, right: 24, color: "#fff", padding: "12px 18px", borderRadius: 10, display: "flex", alignItems: "center", gap: 9, fontSize: 15.5, fontWeight: 600, boxShadow: "0 14px 34px rgba(0,0,0,0.25)", zIndex: 100 },
  storageWarn: { position: "fixed", top: 0, left: 0, right: 0, background: "#9a7b0a", color: "#fff", padding: "8px 16px", fontSize: 14.5, display: "flex", alignItems: "center", gap: 8, justifyContent: "center", zIndex: 200 },
};
