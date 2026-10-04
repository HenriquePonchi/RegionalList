// Lista padrão extraída da planilha "Check List Ronda C. Cirúrgico.xlsx" (96 equipamentos)
const checklistPadrao = [
    { patrimonio: "3133", descricao: "APARELHO DE ANESTESIA", fabricante: "GE", modelo: "CARESTATION 620", serie: "SM 618150017WA", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "3134", descricao: "APARELHO DE ANESTESIA", fabricante: "GE", modelo: "CARESTATION 620", serie: "SM 618150020WA", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "49269", descricao: "APARELHO DE ANESTESIA", fabricante: "MINDRAY", modelo: "WATO EX-30", serie: "DY-19001100", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "87779", descricao: "APARELHO DE ANESTESIA", fabricante: "MINDRAY", modelo: "WATO EX-30", serie: "DY-08002986", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "87780", descricao: "APARELHO DE ANESTESIA", fabricante: "MINDRAY", modelo: "WATO EX-30", serie: "DY-0B002987", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "87781", descricao: "APARELHO DE ANESTESIA", fabricante: "MINDRAY", modelo: "WATO EX-30", serie: "DY-08002985", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2111", descricao: "ASPIRADOR CIRÚRGICO", fabricante: "FANEM", modelo: "DAB.089 AME", serie: "RAE36059", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2772", descricao: "ASPIRADOR CIRÚRGICO", fabricante: "STRYKER", modelo: "5100-50A", serie: "42620", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "3179", descricao: "BAIR HUGGER", fabricante: "3M", modelo: "750", serie: "58323", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "28717", descricao: "BALANÇA ELETRÔNICA", fabricante: "WELMY", modelo: "R/I 109-E", serie: "6023", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "269", descricao: "BERÇO AQUECIDO", fabricante: "FANEM", modelo: "AMPLA", serie: "HAN 056638", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "13404", descricao: "BERÇO AQUECIDO", fabricante: "FANEM", modelo: "AMPLA 2085", serie: "DAB8429", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2792", descricao: "BISTURI ELÉTRICO", fabricante: "WEM", modelo: "SS-501S", serie: "AAB0008088", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2793", descricao: "BISTURI ELÉTRICO", fabricante: "WEM", modelo: "SS-501S", serie: "AAB0008089", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2795", descricao: "BISTURI ELÉTRICO", fabricante: "WEM", modelo: "SS-501S", serie: "AAB0008092", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2797", descricao: "BISTURI ELÉTRICO", fabricante: "WEM", modelo: "SS-501S", serie: "AAB0008087", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2798", descricao: "BISTURI ELÉTRICO", fabricante: "WEM", modelo: "SS-501S", serie: "AAB0008090", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2799", descricao: "BISTURI ELÉTRICO", fabricante: "WEM", modelo: "SS-501S", serie: "AAB0008085", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "10004", descricao: "BISTURI ELÉTRICO", fabricante: "WEM", modelo: "SS601MCA", serie: "656", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "", descricao: "BOMBA INFUSORA SERINGA", fabricante: "SAMTRONIC", modelo: "ST-7000 TCI", serie: "", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "", descricao: "BOMBA INFUSORA SERINGA", fabricante: "SAMTRONIC", modelo: "ST-7000 TCI", serie: "", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "", descricao: "BOMBA INFUSORA SERINGA", fabricante: "SAMTRONIC", modelo: "ST-7000 TCI", serie: "", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "", descricao: "BOMBA INFUSORA SERINGA", fabricante: "SAMTRONIC", modelo: "ST-7000 TCI", serie: "", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "", descricao: "BOMBA INFUSORA VOLUMÉTRICA", fabricante: "SAMTRONIC", modelo: "ST-1000", serie: "", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "", descricao: "BOMBA INFUSORA VOLUMÉTRICA", fabricante: "SAMTRONIC", modelo: "ST-1000", serie: "", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "", descricao: "BOMBA INFUSORA VOLUMÉTRICA", fabricante: "SAMTRONIC", modelo: "ST-1000", serie: "", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "", descricao: "BOMBA INFUSORA VOLUMÉTRICA", fabricante: "SAMTRONIC", modelo: "ST-1000", serie: "", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2735", descricao: "BOMBA PARA HISTEROSCOPIA", fabricante: "STRYKER", modelo: "FLUID SAFE", serie: "1308CE490", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "3222", descricao: "CAIXA AQUECEDORA SORO", fabricante: "FANEM", modelo: "HEATER", serie: "YAP074309", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "3224", descricao: "CAIXA AQUECEDORA SORO", fabricante: "FANEM", modelo: "HEATER", serie: "YAP074310", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2777", descricao: "CÂMERA DE VÍDEO", fabricante: "STRYKER", modelo: "ÓTICA", serie: "17F505214", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2785", descricao: "CÂMERA DE VÍDEO", fabricante: "STRYKER", modelo: "ÓTICA", serie: "17G521884", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "202", descricao: "CARDIOVERSOR", fabricante: "PHILIPS", modelo: "DFM 100", serie: "CN 32610441", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2738", descricao: "CONSOLE CÂMERA DE VIDEOCIRURGIA", fabricante: "STRYKER", modelo: "1488", serie: "17F556694", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2784", descricao: "CONSOLE CÂMERA DE VIDEOCIRURGIA", fabricante: "STRYKER", modelo: "1488", serie: "17F556694", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2718", descricao: "DETECTOR FETAL", fabricante: "MEDPEJ", modelo: "DF7001D", serie: "413197", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "10010", descricao: "EQUIPAMENTO ELETROESTIMULADOR", fabricante: "BBRAUM", modelo: "STIMUPLEX", serie: "106045", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "10018", descricao: "EQUIPAMENTO ELETROESTIMULADOR", fabricante: "BBRAUM", modelo: "STIMUPLEX", serie: "42011051", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "1358", descricao: "FOCO CIRÚRGICO MOVEL", fabricante: "MEDPEJ", modelo: "FL-2000AL", serie: "153168", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "257", descricao: "FOCO CIRÚRGICO TETO", fabricante: "SISMATEC", modelo: "3LE/3LE", serie: "1216724504", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "259", descricao: "FOCO CIRÚRGICO TETO", fabricante: "SISMATEC", modelo: "3LE/3LE", serie: "1216724505", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "260", descricao: "FOCO CIRÚRGICO TETO", fabricante: "SISMATEC", modelo: "3LE/3LE", serie: "1216724503", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "261", descricao: "FOCO CIRÚRGICO TETO", fabricante: "SISMATEC", modelo: "3LE/3LE", serie: "1216724502", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "262", descricao: "FOCO CIRÚRGICO TETO", fabricante: "SISMATEC", modelo: "3LE/3LE", serie: "1216724501", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2737", descricao: "FONTE DE LUZ", fabricante: "STRYKER", modelo: "LED 9000", serie: "17C544704", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2783", descricao: "FONTE DE LUZ", fabricante: "STRYKER", modelo: "LED 9000", serie: "17C544694", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "433", descricao: "INCUBADORA DE TRANSPORTE", fabricante: "FANEM", modelo: "IT -158/TS", serie: "CF 9440", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2736", descricao: "INSUFLADOR VIDEOLAPAROSCOPIA", fabricante: "STRYKER", modelo: "PNEUMOSURE XL", serie: "1705CE0514", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2782", descricao: "INSUFLADOR VIDEOLAPAROSCOPIA", fabricante: "STRYKER", modelo: "PNEUMOSURE XL", serie: "1705CE0517", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "176", descricao: "MESA CIRÚRGICA", fabricante: "BARRFAB", modelo: "BF-683 TDP", serie: "010-3261-116", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "239", descricao: "MESA CIRÚRGICA", fabricante: "BARRFAB", modelo: "BF 683T", serie: "010-3251-116", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "246", descricao: "MESA CIRÚRGICA", fabricante: "BARRFAB", modelo: "BF-683 TDP", serie: "010-3221-116", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "251", descricao: "MESA CIRÚRGICA", fabricante: "BARRFAB", modelo: "BF-683 TDP", serie: "010-3231-116", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "274", descricao: "MESA CIRÚRGICA", fabricante: "BARRFAB", modelo: "BF-683 TDP", serie: "010-3241-116", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "12783", descricao: "MESA CIRÚRGICA", fabricante: "BARRFAB", modelo: "BF-683 TDP", serie: "4886", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "186", descricao: "MÓDULO ANALIZADOR DE GASES", fabricante: "DIXTAL", modelo: "AGOI", serie: "170801663", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "187", descricao: "MÓDULO ANALIZADOR DE GASES", fabricante: "DIXTAL", modelo: "AGOI", serie: "170801659", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "267", descricao: "MÓDULO ANALIZADOR DE GASES", fabricante: "DIXTAL", modelo: "AGOI", serie: "170801665", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "267", descricao: "MÓDULO ANALIZADOR DE GASES", fabricante: "DIXTAL", modelo: "AGOI", serie: "170801665", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "271", descricao: "MÓDULO ANALIZADOR DE GASES", fabricante: "DIXTAL", modelo: "AGOI", serie: "170801664", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "3517", descricao: "MÓDULO DE CAPNOGRAFIA", fabricante: "DIXTAL", modelo: "ETCO2", serie: "171904110", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "3521", descricao: "MÓDULO DE PRESSÃO INVASIVA", fabricante: "DIXTAL", modelo: "IBP", serie: "160L17123", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2740", descricao: "MONITOR DE VÍDEO", fabricante: "STRYKER", modelo: "VISIONPRO 26", serie: "VPD266L0285", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2781", descricao: "MONITOR DE VÍDEO", fabricante: "STRYKER", modelo: "VISIONPRO 26", serie: "VPD267C0414", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2640", descricao: "MONITOR DE VÍDEO NDS", fabricante: "MEDICAL DEVICE", modelo: "ENDOVUE 21", serie: "C16-004551", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "91", descricao: "MONITOR MULTIPARÂMETROS", fabricante: "DIXTAL", modelo: "DX2021", serie: "171509656", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "97", descricao: "MONITOR MULTIPARÂMETROS", fabricante: "DIXTAL", modelo: "DX2021", serie: "171509696", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "99", descricao: "MONITOR MULTIPARÂMETROS", fabricante: "DIXTAL", modelo: "DX2021", serie: "171509713", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "149", descricao: "MONITOR MULTIPARÂMETROS", fabricante: "DIXTAL", modelo: "DX2021", serie: "171509700", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "321", descricao: "MONITOR MULTIPARÂMETROS", fabricante: "DIXTAL", modelo: "DX2021", serie: "171509674", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2580", descricao: "MONITOR MULTIPARÂMETROS", fabricante: "DIXTAL", modelo: "DX2021", serie: "171509687", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2877", descricao: "MONITOR MULTIPARÂMETROS", fabricante: "DIXTAL", modelo: "DX2021", serie: "171311520", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "3110", descricao: "MONITOR MULTIPARÂMETROS", fabricante: "DIXTAL", modelo: "DX2021", serie: "171509670", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "4", descricao: "MONITOR MULTIPARÂMETROS", fabricante: "DIXTAL", modelo: "DX2020", serie: "171311540", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "265", descricao: "MONITOR MULTIPARÂMETROS", fabricante: "DIXTAL", modelo: "DX2020", serie: "171311590", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "270", descricao: "MONITOR MULTIPARÂMETROS", fabricante: "DIXTAL", modelo: "DX2020", serie: "171311570", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2876", descricao: "MONITOR MULTIPARÂMETROS", fabricante: "DIXTAL", modelo: "DX2020", serie: "171311526", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "3113", descricao: "MONITOR MULTIPARÂMETROS", fabricante: "DIXTAL", modelo: "DX2020", serie: "171311512", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2634", descricao: "PROCESSADORA ENDOSCOPIA", fabricante: "FUJIFILM", modelo: "EPX-2500", serie: "3V456G052", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "3327", descricao: "NOBREAK", fabricante: "SMS", modelo: "STATION II", serie: "273920287556", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "", descricao: "RACK C/RODIZIO", fabricante: "LABORMED", modelo: "5 DIVISÓRIAS", serie: "N/C", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2739", descricao: "RACK PARA SISTEMA DE VIDEO", fabricante: "STRYKER", modelo: "5 DIVISÓRIAS", serie: "52753235", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2780", descricao: "RACK PARA SISTEMA DE VIDEO", fabricante: "STRYKER", modelo: "5 DIVISÓRIAS", serie: "527546343", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2679", descricao: "RAIO-X ARCO EM C", fabricante: "GE", modelo: "BRIVO - OEC", serie: "B3SS1700010", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "4184", descricao: "RAIO-X ARCO EM C", fabricante: "PHILIPS/VNI", modelo: "CENOS PLUS", serie: "1750020141", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "12784", descricao: "RAIO-X ARCO EM C", fabricante: "GE", modelo: "EVERVIEW 7500 GE", serie: "OECF2S08208", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "4100", descricao: "REANIMADOR INFANTIL", fabricante: "FANEM", modelo: "BABYPUFF 1020", serie: "FAS097843", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "4102", descricao: "REANIMADOR INFANTIL", fabricante: "FANEM", modelo: "BABYPUFF 1020", serie: "FAS097841", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "12781", descricao: "ULTRASSOM", fabricante: "FUJIFILM/SONOSITE", modelo: "MICROMAXX", serie: "03DX22", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "3052", descricao: "VAPORIZADOR CALIBRADO", fabricante: "PENLON", modelo: "SEVOFLURANE", serie: "D0214-0352", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "3136", descricao: "VAPORIZADOR CALIBRADO", fabricante: "DATEXOHMEDA", modelo: "SEVOFLURANE", serie: "BEJW02053", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "3188", descricao: "VAPORIZADOR CALIBRADO", fabricante: "DATEXOHMEDA", modelo: "TEC7 ISSO", serie: "BEGX1224", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "87783", descricao: "VAPORIZADOR CALIBRADO", fabricante: "MINDRAY", modelo: "SEVOFLURANE-V60", serie: "DK509245111", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "87784", descricao: "VAPORIZADOR CALIBRADO SERVO", fabricante: "MINDRAY", modelo: "V-60", serie: "DK509245097", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "2654", descricao: "VIDEOBRONCOSCOPIO", fabricante: "FUJIFILM", modelo: "EB-530H", serie: "1B087K265", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" },
    { patrimonio: "210", descricao: "VENTILADOR DE TRANSPORTE", fabricante: "DRAGER", modelo: "OXYLOG 3000 PLUS", serie: "ASKF-0111", inspecaoVisual: "", autoTeste: "", inspecaoAcessorios: "", ligadoRede: "" }
];

'use strict';

/* =====================================================================
   CHECKLIST DIGITAL - CENTRO CIRÚRGICO
   HTML5 + CSS3 + JavaScript puro (jsPDF + AutoTable via CDN para o PDF)
   ===================================================================== */

const CHAVE_STORAGE = 'checklist';

const CAMPOS_AVALIACAO = [
    { chave: 'inspecaoVisual',     titulo: 'INSPEÇÃO VISUAL' },
    { chave: 'autoTeste',          titulo: 'AUTO TESTE' },
    { chave: 'inspecaoAcessorios', titulo: 'INSPEÇÃO ACESSÓRIOS' },
    { chave: 'ligadoRede',         titulo: 'LIGADO NA REDE ELÉTRICA' }
];

const VALORES = [
    { v: 'C',  nome: 'Conforme' },
    { v: 'NC', nome: 'Não Conforme' },
    { v: 'NA', nome: 'Não Aplicado' }
];

let estado = null;

/* ---------------------------------------------------------------------
   Utilidades
   --------------------------------------------------------------------- */
const $ = (id) => document.getElementById(id);

function dataHojeISO() {
    const d = new Date();
    const p = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function formatarData(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    return m ? `${m[3]}/${m[2]}/${m[1]}` : '__/__/____';
}

function formatarHora(hhmm) {
    return /^\d{2}:\d{2}$/.test(hhmm || '') ? hhmm : '__:__';
}

function valorValido(v) {
    return v === 'C' || v === 'NC' || v === 'NA' ? v : '';
}

/* ---------------------------------------------------------------------
   Estado / armazenamento
   --------------------------------------------------------------------- */
function criarEstadoBase(equipamentos) {
    return {
        versao: 1,
        equipamentos: equipamentos,
        cabecalho: { unidade: '', inicio: '', termino: '', data: dataHojeISO() },
        observacoes: '',
        tecnico: { nome: '', registro: '', data: dataHojeISO() },
        enfermeiro: { nome: '', coren: '' },
        assinaturas: { tecnico: '', enfermeiro: '' }
    };
}

// Cópia nova da lista padrão da planilha (ids estáveis: p001, p002, ...)
function clonarListaPadrao() {
    return checklistPadrao.map((e, i) => ({
        ...e,
        id: 'p' + String(i + 1).padStart(3, '0'),
        padrao: true,
        inspecaoVisual: '',
        autoTeste: '',
        inspecaoAcessorios: '',
        ligadoRede: ''
    }));
}

function carregarChecklistPadrao() {
    estado = criarEstadoBase(clonarListaPadrao());
    salvarChecklist();
}

function carregarChecklist() {
    let bruto = null;
    try { bruto = localStorage.getItem(CHAVE_STORAGE); } catch (e) { bruto = null; }

    if (!bruto) {
        carregarChecklistPadrao();
        return;
    }

    try {
        const salvo = JSON.parse(bruto);
        if (!salvo || !Array.isArray(salvo.equipamentos)) throw new Error('formato inválido');

        const base = criarEstadoBase([]);
        estado = {
            ...base,
            ...salvo,
            cabecalho:   { ...base.cabecalho,   ...(salvo.cabecalho || {}) },
            tecnico:     { ...base.tecnico,     ...(salvo.tecnico || {}) },
            enfermeiro:  { ...base.enfermeiro,  ...(salvo.enfermeiro || {}) },
            assinaturas: { ...base.assinaturas, ...(salvo.assinaturas || {}) }
        };
        estado.equipamentos = salvo.equipamentos.map((e, i) => {
            const eq = {
                id: String(e.id || 'n' + Date.now() + '_' + i),
                padrao: !!e.padrao,
                patrimonio: String(e.patrimonio ?? ''),
                descricao: String(e.descricao ?? ''),
                fabricante: String(e.fabricante ?? ''),
                modelo: String(e.modelo ?? ''),
                serie: String(e.serie ?? '')
            };
            CAMPOS_AVALIACAO.forEach((c) => { eq[c.chave] = valorValido(e[c.chave]); });
            return eq;
        });
    } catch (erro) {
        console.warn('Dados salvos inválidos; carregando lista padrão.', erro);
        carregarChecklistPadrao();
    }
}

function salvarChecklist() {
    const indicador = $('statusSalvo');
    try {
        localStorage.setItem(CHAVE_STORAGE, JSON.stringify(estado));
        if (indicador) {
            const h = new Date();
            const p = (n) => String(n).padStart(2, '0');
            indicador.textContent = `Salvo automaticamente às ${p(h.getHours())}:${p(h.getMinutes())}:${p(h.getSeconds())}`;
        }
    } catch (erro) {
        console.error('Falha ao salvar no localStorage', erro);
        if (indicador) indicador.textContent = 'Não foi possível salvar (armazenamento indisponível)';
    }
}

/* ---------------------------------------------------------------------
   Diálogos
   --------------------------------------------------------------------- */
function confirmar({ titulo, mensagem, detalhe = '', ok = 'Confirmar', cancelar = 'Cancelar', semCancelar = false }) {
    return new Promise((resolve) => {
        const dlg = $('dlgConfirmar');
        $('confTitulo').textContent = titulo;
        $('confMensagem').textContent = mensagem;
        $('confDetalhe').textContent = detalhe;
        $('confOk').textContent = ok;
        $('confCancelar').textContent = cancelar;
        $('confCancelar').hidden = semCancelar;
        dlg.returnValue = 'cancel';
        dlg.addEventListener('close', () => resolve(dlg.returnValue === 'ok'), { once: true });
        dlg.showModal();
    });
}

function avisar(titulo, mensagem) {
    return confirmar({ titulo, mensagem, ok: 'OK', semCancelar: true });
}

/* ---------------------------------------------------------------------
   Tabela
   --------------------------------------------------------------------- */
function criarLinha(eq) {
    const tr = document.createElement('tr');
    tr.dataset.id = eq.id;

    [eq.patrimonio, eq.descricao, eq.fabricante, eq.modelo, eq.serie].forEach((texto) => {
        const td = document.createElement('td');
        td.textContent = texto;
        tr.appendChild(td);
    });

    CAMPOS_AVALIACAO.forEach((campo) => {
        const td = document.createElement('td');
        td.className = 'cel-status';
        td.dataset.campo = campo.chave;

        const grupo = document.createElement('div');
        grupo.className = 'seg';
        grupo.setAttribute('role', 'group');
        grupo.setAttribute('aria-label', campo.titulo);

        VALORES.forEach((op) => {
            const b = document.createElement('button');
            b.type = 'button';
            b.className = 'opcao';
            b.dataset.v = op.v;
            b.textContent = op.v;
            b.title = `${op.nome} - ${campo.titulo}`;
            b.setAttribute('aria-label', `${campo.titulo}: ${op.nome}`);
            b.setAttribute('aria-pressed', eq[campo.chave] === op.v ? 'true' : 'false');
            grupo.appendChild(b);
        });

        td.appendChild(grupo);
        td.classList.toggle('nc', eq[campo.chave] === 'NC');
        tr.appendChild(td);
    });

    const tdAcao = document.createElement('td');
    tdAcao.style.textAlign = 'center';
    const bExcluir = document.createElement('button');
    bExcluir.type = 'button';
    bExcluir.className = 'btn-excluir';
    bExcluir.dataset.acao = 'excluir';
    bExcluir.textContent = 'Excluir';
    bExcluir.title = 'Remover este equipamento';
    tdAcao.appendChild(bExcluir);
    tr.appendChild(tdAcao);

    return tr;
}

function renderizarTabela() {
    const corpo = $('corpoTabela');
    corpo.textContent = '';

    if (estado.equipamentos.length === 0) {
        const tr = document.createElement('tr');
        const td = document.createElement('td');
        td.colSpan = 10;
        td.className = 'vazio';
        td.textContent = 'Nenhum equipamento na lista. Use "+ ADICIONAR EQUIPAMENTO" ou "RESTAURAR PADRÃO".';
        tr.appendChild(td);
        corpo.appendChild(tr);
        return;
    }

    const frag = document.createDocumentFragment();
    estado.equipamentos.forEach((eq) => frag.appendChild(criarLinha(eq)));
    corpo.appendChild(frag);
}

function selecionarStatus(id, campo, valor) {
    const eq = estado.equipamentos.find((e) => e.id === id);
    if (!eq || !CAMPOS_AVALIACAO.some((c) => c.chave === campo) || !valorValido(valor)) return;

    // Clicar de novo na opção já marcada desmarca (permite corrigir engano)
    eq[campo] = eq[campo] === valor ? '' : valor;

    const tr = $('corpoTabela').querySelector(`tr[data-id="${CSS.escape(id)}"]`);
    if (tr) {
        const td = tr.querySelector(`td[data-campo="${campo}"]`);
        td.querySelectorAll('.opcao').forEach((b) => {
            b.setAttribute('aria-pressed', b.dataset.v === eq[campo] ? 'true' : 'false');
        });
        td.classList.toggle('nc', eq[campo] === 'NC');
    }

    salvarChecklist();
    atualizarResumo();
}

function atualizarResumo() {
    let c = 0, nc = 0, na = 0, pend = 0;
    estado.equipamentos.forEach((eq) => {
        CAMPOS_AVALIACAO.forEach((campo) => {
            const v = eq[campo.chave];
            if (v === 'C') c++;
            else if (v === 'NC') nc++;
            else if (v === 'NA') na++;
            else pend++;
        });
    });
    $('resTotal').textContent = estado.equipamentos.length;
    $('resC').textContent = c;
    $('resNC').textContent = nc;
    $('resNA').textContent = na;
    $('resPend').textContent = pend;
    $('resItens').textContent = c + nc + na + pend;
}

/* ---------------------------------------------------------------------
   Equipamentos: adicionar / remover / restaurar / novo checklist
   --------------------------------------------------------------------- */
function adicionarEquipamento(dados) {
    const eq = {
        id: 'n' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
        padrao: false,
        patrimonio: (dados.patrimonio || '').trim(),
        descricao: (dados.descricao || '').trim(),
        fabricante: (dados.fabricante || '').trim(),
        modelo: (dados.modelo || '').trim(),
        serie: (dados.serie || '').trim(),
        inspecaoVisual: '',
        autoTeste: '',
        inspecaoAcessorios: '',
        ligadoRede: ''
    };
    estado.equipamentos.push(eq);
    renderizarTabela();
    salvarChecklist();
    atualizarResumo();

    const tr = $('corpoTabela').querySelector(`tr[data-id="${CSS.escape(eq.id)}"]`);
    if (tr) {
        tr.classList.add('destaque');
        tr.scrollIntoView({ block: 'center', behavior: 'smooth' });
    }
}

async function removerEquipamento(id) {
    const ok = await confirmar({
        titulo: 'Remover equipamento',
        mensagem: 'Tem certeza que deseja remover este equipamento?',
        ok: 'Remover',
        cancelar: 'Cancelar'
    });
    if (!ok) return;
    estado.equipamentos = estado.equipamentos.filter((e) => e.id !== id);
    renderizarTabela();
    salvarChecklist();
    atualizarResumo();
}

async function restaurarPadrao() {
    const ok = await confirmar({
        titulo: 'Restaurar lista padrão',
        mensagem: 'A lista de equipamentos voltará exatamente à da planilha original: itens removidos retornam e itens adicionados por você são excluídos.',
        detalhe: 'As respostas C/NC/NA dos equipamentos originais que continuam na lista, as observações, os dados e as assinaturas são mantidos.',
        ok: 'Restaurar padrão'
    });
    if (!ok) return;

    const atuais = new Map(estado.equipamentos.filter((e) => e.padrao).map((e) => [e.id, e]));
    estado.equipamentos = clonarListaPadrao().map((novo) => {
        const anterior = atuais.get(novo.id);
        if (anterior) CAMPOS_AVALIACAO.forEach((c) => { novo[c.chave] = anterior[c.chave]; });
        return novo;
    });
    renderizarTabela();
    salvarChecklist();
    atualizarResumo();
}

async function limparChecklist() {
    const ok = await confirmar({
        titulo: 'Novo checklist',
        mensagem: 'Deseja limpar o checklist atual? Serão apagados: respostas C/NC/NA, horários, data, observações, dados do técnico e do enfermeiro e as assinaturas.',
        detalhe: 'A lista de equipamentos (incluindo os adicionados) e a Unidade/Setor são mantidas.',
        ok: 'Limpar e começar novo'
    });
    if (!ok) return;

    const base = criarEstadoBase(estado.equipamentos);
    estado.equipamentos.forEach((eq) => CAMPOS_AVALIACAO.forEach((c) => { eq[c.chave] = ''; }));
    estado.cabecalho = { ...base.cabecalho, unidade: estado.cabecalho.unidade };
    estado.observacoes = '';
    estado.tecnico = base.tecnico;
    estado.enfermeiro = base.enfermeiro;

    limparAssinatura('tecnico', false);
    limparAssinatura('enfermeiro', false);

    aplicarEstadoNosCampos();
    renderizarTabela();
    salvarChecklist();
    atualizarResumo();
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

/* ---------------------------------------------------------------------
   Campos de texto ligados ao estado (data-bind="a.b")
   --------------------------------------------------------------------- */
function lerCaminho(caminho) {
    return caminho.split('.').reduce((o, k) => (o == null ? o : o[k]), estado);
}

function gravarCaminho(caminho, valor) {
    const partes = caminho.split('.');
    const ultimo = partes.pop();
    const alvo = partes.reduce((o, k) => o[k], estado);
    alvo[ultimo] = valor;
}

function aplicarEstadoNosCampos() {
    document.querySelectorAll('[data-bind]').forEach((el) => {
        el.value = lerCaminho(el.dataset.bind) ?? '';
    });
}

function configurarCampos() {
    document.querySelectorAll('[data-bind]').forEach((el) => {
        el.addEventListener('input', () => {
            gravarCaminho(el.dataset.bind, el.value);
            salvarChecklist();
        });
    });
}

/* ---------------------------------------------------------------------
   Assinaturas (Canvas + Pointer Events: mouse, toque e caneta)
   --------------------------------------------------------------------- */
const assinaturas = {};

function configurarAssinatura(chave, idCanvas, idCaixa, idBotao) {
    const canvas = $(idCanvas);
    const info = {
        chave,
        canvas,
        caixa: $(idCaixa),
        ctx: canvas.getContext('2d'),
        largura: 0,
        altura: 0,
        desenhando: false,
        ultimo: null,
        meio: null
    };
    assinaturas[chave] = info;

    ajustarCanvas(info);

    if ('ResizeObserver' in window) {
        new ResizeObserver(() => ajustarCanvas(info)).observe(canvas);
    } else {
        window.addEventListener('resize', () => ajustarCanvas(info));
    }

    const posicao = (ev) => {
        const r = canvas.getBoundingClientRect();
        return { x: ev.clientX - r.left, y: ev.clientY - r.top };
    };
    const largura = (ev) =>
        ev.pointerType === 'pen' && ev.pressure > 0 ? 1.2 + ev.pressure * 2.6 : 2.4;

    canvas.addEventListener('pointerdown', (ev) => {
        if (ev.pointerType === 'mouse' && ev.button !== 0) return;
        ev.preventDefault();
        try { canvas.setPointerCapture(ev.pointerId); } catch (e) { /* ignora */ }
        const p = posicao(ev);
        info.desenhando = true;
        info.ultimo = p;
        info.meio = p;
        info.ctx.fillStyle = info.ctx.strokeStyle;
        info.ctx.beginPath();
        info.ctx.arc(p.x, p.y, largura(ev) / 2, 0, Math.PI * 2);
        info.ctx.fill();
        info.caixa.classList.add('com-tinta');
    });

    canvas.addEventListener('pointermove', (ev) => {
        if (!info.desenhando) return;
        ev.preventDefault();
        const eventos = typeof ev.getCoalescedEvents === 'function' && ev.getCoalescedEvents().length
            ? ev.getCoalescedEvents()
            : [ev];
        eventos.forEach((e) => {
            const p = posicao(e);
            const mid = { x: (info.ultimo.x + p.x) / 2, y: (info.ultimo.y + p.y) / 2 };
            info.ctx.lineWidth = largura(e);
            info.ctx.beginPath();
            info.ctx.moveTo(info.meio.x, info.meio.y);
            info.ctx.quadraticCurveTo(info.ultimo.x, info.ultimo.y, mid.x, mid.y);
            info.ctx.stroke();
            info.ultimo = p;
            info.meio = mid;
        });
    });

    const finalizar = (ev) => {
        if (!info.desenhando) return;
        info.desenhando = false;
        try { canvas.releasePointerCapture(ev.pointerId); } catch (e) { /* ignora */ }
        salvarAssinatura(info);
    };
    canvas.addEventListener('pointerup', finalizar);
    canvas.addEventListener('pointercancel', finalizar);
    canvas.addEventListener('contextmenu', (ev) => ev.preventDefault());

    $(idBotao).addEventListener('click', () => limparAssinatura(chave, true));
}

function configurarAssinaturaTecnico() {
    configurarAssinatura('tecnico', 'canvasTecnico', 'boxAssTecnico', 'btnLimparTecnico');
}

function configurarAssinaturaEnfermeiro() {
    configurarAssinatura('enfermeiro', 'canvasEnfermeiro', 'boxAssEnfermeiro', 'btnLimparEnfermeiro');
}

// Dimensiona o canvas (nítido em telas retina) e redesenha a assinatura salva
function ajustarCanvas(info) {
    const r = info.canvas.getBoundingClientRect();
    const w = Math.round(r.width);
    const h = Math.round(r.height);
    if (w === 0 || h === 0) return;
    if (w === info.largura && h === info.altura) return;

    const dpr = Math.max(1, window.devicePixelRatio || 1);
    info.largura = w;
    info.altura = h;
    info.canvas.width = Math.round(w * dpr);
    info.canvas.height = Math.round(h * dpr);

    const ctx = info.ctx;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.strokeStyle = '#10233a';
    ctx.lineWidth = 2.4;

    const salva = estado && estado.assinaturas[info.chave];
    info.caixa.classList.toggle('com-tinta', !!salva);
    if (salva) {
        const img = new Image();
        img.onload = () => {
            const razaoImg = img.width / img.height;
            const razaoCanvas = w / h;
            const dw = razaoImg > razaoCanvas ? w : h * razaoImg;
            const dh = razaoImg > razaoCanvas ? w / razaoImg : h;
            ctx.clearRect(0, 0, w, h);
            ctx.drawImage(img, (w - dw) / 2, (h - dh) / 2, dw, dh);
        };
        img.src = salva;
    }
}

function salvarAssinatura(info) {
    estado.assinaturas[info.chave] = info.canvas.toDataURL('image/png');
    salvarChecklist();
}

function limparAssinatura(chave, salvar = true) {
    const info = assinaturas[chave];
    if (!info) return;
    info.ctx.clearRect(0, 0, info.largura, info.altura);
    info.caixa.classList.remove('com-tinta');
    estado.assinaturas[chave] = '';
    if (salvar) salvarChecklist();
}

/* ---------------------------------------------------------------------
   Exportação para PDF
   --------------------------------------------------------------------- */
async function exportarPDF() {
    if (!window.jspdf || !window.jspdf.jsPDF) {
        await avisar(
            'Biblioteca de PDF indisponível',
            'Não foi possível carregar a biblioteca jsPDF. Verifique a conexão com a internet e recarregue a página.'
        );
        return;
    }

    let pendentes = 0;
    estado.equipamentos.forEach((eq) => CAMPOS_AVALIACAO.forEach((c) => { if (!eq[c.chave]) pendentes++; }));

    if (pendentes > 0) {
        const continuar = await confirmar({
            titulo: 'Itens pendentes',
            mensagem: 'Existem itens ainda não avaliados. Deseja continuar com a exportação?',
            detalhe: `${pendentes} item(ns) de inspeção sem avaliação (aparecerão como "-" no PDF).`,
            ok: 'Continuar',
            cancelar: 'Voltar ao checklist'
        });
        if (!continuar) return;
    }

    try {
        const { doc } = escolherMelhorPDF();
        const nome = `Checklist_Centro_Cirurgico_${estado.cabecalho.data || dataHojeISO()}.pdf`;
        doc.save(nome);
    } catch (erro) {
        console.error(erro);
        await avisar('Erro ao gerar o PDF', 'Ocorreu um erro ao gerar o PDF: ' + (erro && erro.message ? erro.message : erro));
    }
}

// Testa tamanhos de fonte do maior para o menor (mínimo legível 6,5 pt) e fica
// com o que gera MENOS páginas; em caso de empate, a fonte maior.
function escolherMelhorPDF() {
    const candidatos = [
        { fs: 8,   pad: 0.8 },
        { fs: 7.5, pad: 0.7 },
        { fs: 7,   pad: 0.6 },
        { fs: 6.5, pad: 0.5 }
    ];
    let melhor = null;
    for (const cfg of candidatos) {
        const doc = criarPDF(cfg);
        const paginas = doc.getNumberOfPages();
        if (!melhor || paginas < melhor.paginas) melhor = { doc, paginas, cfg };
    }
    return melhor;
}

function criarPDF(cfg) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });

    const PW = doc.internal.pageSize.getWidth();   // 297
    const PH = doc.internal.pageSize.getHeight();  // 210
    const M = { l: 7, r: 7, t: 7, b: 10 };
    const largura = PW - M.l - M.r;
    const topoConteudo = M.t + 13.5;
    const limiteY = PH - M.b;
    const cab = estado.cabecalho;

    // Cabeçalho repetido em todas as páginas
    const desenharCabecalho = () => {
        doc.setTextColor(0);
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(13);
        doc.text('CHECKLIST - CENTRO CIRÚRGICO', M.l, M.t + 5);

        doc.setFontSize(9);
        doc.text(
            `INÍCIO: ${formatarHora(cab.inicio)}      TÉRMINO: ${formatarHora(cab.termino)}      DATA: ${formatarData(cab.data)}`,
            PW - M.r, M.t + 5, { align: 'right' }
        );

        if (cab.unidade && cab.unidade.trim()) {
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(8.5);
            doc.text(`Unidade/Setor: ${cab.unidade.trim()}`, M.l, M.t + 9.6);
        }

        doc.setDrawColor(0);
        doc.setLineWidth(0.35);
        doc.line(M.l, M.t + 11.2, PW - M.r, M.t + 11.2);
    };

    // Tabela principal
    const corpo = estado.equipamentos.map((eq) => [
        eq.patrimonio, eq.descricao, eq.fabricante, eq.modelo, eq.serie,
        eq.inspecaoVisual || '-',
        eq.autoTeste || '-',
        eq.inspecaoAcessorios || '-',
        eq.ligadoRede || '-'
    ]);

    doc.autoTable({
        head: [['PATRIM.', 'DESCRIÇÃO', 'FABRICANTE', 'MODELO', 'Nº DE SÉRIE',
                'INSPEÇÃO VISUAL', 'AUTO TESTE', 'INSPEÇÃO ACESSÓRIOS', 'LIGADO NA REDE ELÉTRICA']],
        body: corpo,
        startY: topoConteudo,
        margin: { top: topoConteudo, left: M.l, right: M.r, bottom: M.b },
        theme: 'grid',
        showHead: 'everyPage',
        rowPageBreak: 'avoid',
        styles: {
            font: 'helvetica',
            fontSize: cfg.fs,
            cellPadding: { top: cfg.pad, bottom: cfg.pad, left: 1.2, right: 1.2 },
            lineColor: [70, 70, 70],
            lineWidth: 0.1,
            textColor: 0,
            valign: 'middle',
            overflow: 'linebreak'
        },
        headStyles: {
            fillColor: [222, 222, 222],
            textColor: 0,
            fontStyle: 'bold',
            halign: 'center',
            lineWidth: 0.15
        },
        alternateRowStyles: { fillColor: [246, 246, 246] },
        columnStyles: {
            0: { cellWidth: 16, halign: 'center' },
            1: { cellWidth: 60 },
            2: { cellWidth: 33 },
            3: { cellWidth: 36 },
            4: { cellWidth: 38 },
            5: { cellWidth: 25, halign: 'center', fontStyle: 'bold' },
            6: { cellWidth: 25, halign: 'center', fontStyle: 'bold' },
            7: { cellWidth: 25, halign: 'center', fontStyle: 'bold' },
            8: { cellWidth: 25, halign: 'center', fontStyle: 'bold' }
        },
        // NC: cinza + negrito (legível mesmo impresso em preto e branco)
        didParseCell: (data) => {
            if (data.section === 'body' && data.column.index >= 5 && data.cell.raw === 'NC') {
                data.cell.styles.fillColor = [205, 205, 205];
                data.cell.styles.fontStyle = 'bold';
            }
        },
        didDrawPage: desenharCabecalho
    });

    let y = (doc.lastAutoTable ? doc.lastAutoTable.finalY : topoConteudo) + 4;

    const garantirEspaco = (altura) => {
        if (y + altura > limiteY) {
            doc.addPage();
            desenharCabecalho();
            y = topoConteudo;
        }
    };

    // ---- Observações adicionais ----
    const lhObs = 3.7;
    const texto = (estado.observacoes || '').replace(/\r/g, '').trim();
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8.5);
    const linhasObs = texto ? doc.splitTextToSize(texto, largura - 4) : [];

    garantirEspaco(6 + Math.min(Math.max(linhasObs.length, 3), 3) * lhObs + 3);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9);
    doc.setTextColor(0);
    doc.text('OBSERVAÇÕES ADICIONAIS', M.l, y + 3);
    y += 5;

    let i = 0;
    do {
        const maxLinhas = Math.max(1, Math.floor((limiteY - y - 5) / lhObs));
        const n = Math.min(linhasObs.length - i, maxLinhas);
        const ultimoTrecho = i + n >= linhasObs.length;
        const alturaCaixa = ultimoTrecho ? Math.max(14, n * lhObs + 4) : n * lhObs + 4;

        doc.setDrawColor(0);
        doc.setLineWidth(0.2);
        doc.rect(M.l, y, largura, alturaCaixa);
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(8.5);
        for (let k = 0; k < n; k++) {
            doc.text(linhasObs[i + k], M.l + 2, y + 3.6 + k * lhObs);
        }
        i += n;
        y += alturaCaixa + 3;

        if (i < linhasObs.length) {
            doc.addPage();
            desenharCabecalho();
            y = topoConteudo;
        }
    } while (i < linhasObs.length);

    // ---- Legenda ----
    garantirEspaco(7);
    doc.setFontSize(8.5);
    doc.setFont('helvetica', 'bold');
    doc.text('LEGENDA:', M.l, y + 3);
    doc.setFont('helvetica', 'normal');
    doc.text('C - Conforme', M.l + 20, y + 3);
    doc.text('NC - Não Conforme', M.l + 55, y + 3);
    doc.text('NA - Não Aplicado', M.l + 100, y + 3);
    y += 8;

    // ---- Responsáveis e assinaturas ----
    const alturaBloco = 60;
    garantirEspaco(alturaBloco);

    const gap = 8;
    const colW = (largura - gap) / 2;

    const campoLinha = (x, yy, w, rotulo, valor) => {
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(8.5);
        doc.text(rotulo, x, yy);
        const lw = doc.getTextWidth(rotulo) + 1.5;
        doc.setDrawColor(0);
        doc.setLineWidth(0.15);
        doc.line(x + lw, yy + 1, x + w, yy + 1);
        if (valor) {
            doc.setFont('helvetica', 'bold');
            doc.text(String(valor), x + lw + 1, yy - 0.2, { maxWidth: w - lw - 1 });
        }
    };

    const colunaResponsavel = (x, titulo, rotuloRegistro, registro, nome, dataIso, imagem, rotuloAss) => {
        doc.setFillColor(230, 230, 230);
        doc.setDrawColor(0);
        doc.setLineWidth(0.2);
        doc.rect(x, y, colW, 5.5, 'FD');
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(9);
        doc.text(titulo, x + 2, y + 3.9);

        campoLinha(x, y + 11, colW, 'Nome: ', nome);
        campoLinha(x, y + 17, colW, rotuloRegistro, registro);
        if (dataIso !== null) campoLinha(x, y + 23, colW, 'Data: ', formatarData(dataIso));

        const linhaY = y + alturaBloco - 8;
        if (imagem) {
            try {
                const props = doc.getImageProperties(imagem);
                const razao = props.width / props.height;
                const maxW = colW - 14;
                const maxH = 26;
                const w = Math.min(maxW, maxH * razao);
                const h = w / razao;
                doc.addImage(imagem, 'PNG', x + (colW - w) / 2, linhaY - h - 0.5, w, h);
            } catch (e) {
                console.warn('Não foi possível inserir a assinatura no PDF', e);
            }
        }
        doc.setDrawColor(0);
        doc.setLineWidth(0.3);
        doc.line(x + 8, linhaY, x + colW - 8, linhaY);
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(8.5);
        doc.text(rotuloAss, x + colW / 2, linhaY + 4, { align: 'center' });
    };

    colunaResponsavel(
        M.l, 'TÉCNICO RESPONSÁVEL', 'Registro/Identificação: ', estado.tecnico.registro,
        estado.tecnico.nome, estado.tecnico.data || '', estado.assinaturas.tecnico, 'Assinatura do Técnico'
    );
    colunaResponsavel(
        M.l + colW + gap, 'ENFERMEIRO RESPONSÁVEL', 'Registro/COREN: ', estado.enfermeiro.coren,
        estado.enfermeiro.nome, null, estado.assinaturas.enfermeiro, 'Assinatura do Enfermeiro'
    );

    // ---- Rodapé com numeração em todas as páginas ----
    const total = doc.getNumberOfPages();
    const agora = new Date();
    const p2 = (n) => String(n).padStart(2, '0');
    const carimbo = `Gerado em ${p2(agora.getDate())}/${p2(agora.getMonth() + 1)}/${agora.getFullYear()} ${p2(agora.getHours())}:${p2(agora.getMinutes())}`;
    for (let pg = 1; pg <= total; pg++) {
        doc.setPage(pg);
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(7.5);
        doc.setTextColor(80);
        doc.text(carimbo, M.l, PH - 4.5);
        doc.text(`Página ${pg} de ${total}`, PW - M.r, PH - 4.5, { align: 'right' });
    }
    doc.setTextColor(0);

    return doc;
}

/* ---------------------------------------------------------------------
   Inicialização
   --------------------------------------------------------------------- */
function configurarEventos() {
    $('btnNovo').addEventListener('click', limparChecklist);
    $('btnRestaurar').addEventListener('click', restaurarPadrao);
    $('btnPdf').addEventListener('click', exportarPDF);

    // Delegação de eventos da tabela (C / NC / NA e Excluir)
    $('corpoTabela').addEventListener('click', (ev) => {
        const tr = ev.target.closest('tr[data-id]');
        if (!tr) return;
        const opcao = ev.target.closest('button.opcao');
        if (opcao) {
            const campo = opcao.closest('td').dataset.campo;
            selecionarStatus(tr.dataset.id, campo, opcao.dataset.v);
            return;
        }
        if (ev.target.closest('button[data-acao="excluir"]')) {
            removerEquipamento(tr.dataset.id);
        }
    });

    // Formulário de novo equipamento
    const dlgAdd = $('dlgAdicionar');
    const form = $('formAdicionar');
    $('btnAdicionar').addEventListener('click', () => {
        form.reset();
        $('erroForm').hidden = true;
        dlgAdd.showModal();
        form.elements.patrimonio.focus();
    });
    $('btnCancelarAdd').addEventListener('click', () => dlgAdd.close());
    form.addEventListener('submit', (ev) => {
        ev.preventDefault();
        const d = Object.fromEntries(new FormData(form).entries());
        if (!(d.descricao || '').trim()) {
            $('erroForm').hidden = false;
            form.elements.descricao.focus();
            return;
        }
        adicionarEquipamento(d);
        dlgAdd.close();
    });
}

document.addEventListener('DOMContentLoaded', () => {
    carregarChecklist();
    aplicarEstadoNosCampos();
    configurarCampos();
    renderizarTabela();
    atualizarResumo();
    configurarAssinaturaTecnico();
    configurarAssinaturaEnfermeiro();
    configurarEventos();
});
