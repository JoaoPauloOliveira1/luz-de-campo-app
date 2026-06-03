import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import maplibregl from 'maplibre-gl';
import * as XLSX from 'xlsx';
import { RECIFE_BOUNDARY } from './data/recifeBoundary.js';

const EXPORT_FIELDS = [
  'RPA',
  'LOCALIZACA',
  'ENDERECO',
  'BARRAMENTO',
  'LATITUDE',
  'LONGITUDE',
  'BAIRRO',
  'MEDICAO',
  'TIPO_LUMIN',
  'TIPO_DE_PO',
  'TIPO_LAMPA',
  'QTDE',
  'POTENCIA',
  'PERDAS',
  'TOTAL_CARGA',
  'CONSUMO_kW',
  'CONSUMO_kW_MES',
  'ATUALIZAÇ',
  'DATA_ATUALIZACAO_DEIL',
  'BLOQ_AMP',
  'OPERADOR',
  'ID_CADASTRO',
  'TIPO_IMPLANTACAO',
  'MOTIVO_IMPLANTACAO',
  'OBRA_NOME',
  'IMPLANTACAO_CONCLUIDA',
  'POTENCIA_LUMINARIA_1',
  'POTENCIA_LUMINARIA_2',
  'POTENCIA_LUMINARIA_3',
  'POTENCIA_LUMINARIA_4',
  'POTENCIA_LUMINARIA_5',
  'TEM_IMAGEM',
  'LINKS_IMAGENS',
  'LINK_IMAGEM_1',
  'LINK_IMAGEM_2',
  'LINK_IMAGEM_3',
  'LINK_IMAGEM_4',
  'LINK_IMAGEM_5',
  'SYNCED_EM',
  'EXPORTED_EM',
];

const MANAGER_SUMMARY_FIELDS = [
  'OPERADOR',
  'IMPLANTACAO_CONCLUIDA',
  'SYNCED_EM',
  'EXPORTED_EM',
];

const MANAGER_LOCATION_FIELDS = [
  'RPA',
  'LOCALIZACA',
  'ENDERECO',
  'BARRAMENTO',
  'LATITUDE',
  'LONGITUDE',
  'BAIRRO',
  'MEDICAO',
  'ID_CADASTRO',
];

const MANAGER_INSTALLATION_FIELDS = [
  'TIPO_IMPLANTACAO',
  'MOTIVO_IMPLANTACAO',
  'OBRA_NOME',
  'TIPO_LUMIN',
  'TIPO_DE_PO',
  'TIPO_LAMPA',
  'QTDE',
  'POTENCIA',
  'PERDAS',
  'TOTAL_CARGA',
  'CONSUMO_kW',
  'CONSUMO_kW_MES',
  'ATUALIZAÃ‡',
  'DATA_ATUALIZACAO_DEIL',
  'BLOQ_AMP',
];

const MANAGER_LUMINAIRE_FIELDS = [
  'POTENCIA_LUMINARIA_1',
  'POTENCIA_LUMINARIA_2',
  'POTENCIA_LUMINARIA_3',
  'POTENCIA_LUMINARIA_4',
  'POTENCIA_LUMINARIA_5',
  'TEM_IMAGEM',
  'LINKS_IMAGENS',
  'LINK_IMAGEM_1',
  'LINK_IMAGEM_2',
  'LINK_IMAGEM_3',
  'LINK_IMAGEM_4',
  'LINK_IMAGEM_5',
];

const MANAGER_COMPACT_TABLE_FIELDS = [
  'RPA',
  'BAIRRO',
  'ENDERECO',
  'TIPO_DE_PO',
  'MOTIVO_IMPLANTACAO',
  'OBRA_NOME',
  'IMPLANTACAO_CONCLUIDA',
  'OPERADOR',
  'SYNCED_EM',
];

const MANAGER_DETAIL_ONLY_FIELDS = EXPORT_FIELDS.filter((field) => !MANAGER_COMPACT_TABLE_FIELDS.includes(field));

const MANAGER_EDITABLE_FIELDS = new Set([
  'RPA',
  'ENDERECO',
  'BAIRRO',
  'MOTIVO_IMPLANTACAO',
  'OBRA_NOME',
  'TIPO_DE_PO',
  'IMPLANTACAO_CONCLUIDA',
]);

const INITIAL_FORM = {
  OPERADOR: '',
  ID_CADASTRO: '',
  TIPO_IMPLANTACAO: '',
  MOTIVO_IMPLANTACAO: '',
  OBRA_NOME: '',
  IMPLANTACAO_CONCLUIDA: '',
  IMAGEM: null,
  RPA: '',
  LOCALIZACA: '',
  ENDERECO: '',
  BARRAMENTO: '',
  LATITUDE: '',
  LONGITUDE: '',
  BAIRRO: '',
  MEDICAO: '',
  TIPO_LUMIN: '',
  TIPO_DE_PO: '',
  TIPO_LAMPA: '',
  QTDE: '1',
  LUMINARIAS: [createEmptyLuminaire(1)],
  PERDAS: '',
  TOTAL_CARGA: '',
  CONSUMO_kW: '',
  CONSUMO_kW_MES: '',
  ATUALIZACAO: new Date().toISOString().slice(0, 10),
  BLOQ_AMP: '',
};

const REQUIRED_FIELDS = [];

const FIELD_LABELS = {
  OPERADOR: 'Operador',
  ID_CADASTRO: 'ID cadastro',
  TIPO_IMPLANTACAO: 'Situação do poste',
  MOTIVO_IMPLANTACAO: 'Motivo de implantação',
  OBRA_NOME: 'Obra',
  IMPLANTACAO_CONCLUIDA: 'Implantação concluída',
  RPA: 'RPA',
  LOCALIZACA: 'Localização',
  ENDERECO: 'Endereço',
  BARRAMENTO: 'Barramento',
  LATITUDE: 'Latitude',
  LONGITUDE: 'Longitude',
  BAIRRO: 'Bairro',
  MEDICAO: 'Medição',
  TIPO_LUMIN: 'Tipo de luminária',
  TIPO_DE_PO: 'Tipo de poste',
  TIPO_LAMPA: 'Tipo de lâmpada',
  QTDE: 'Quantidade',
  POTENCIA: 'Potência (W)',
  PERDAS: 'Perdas',
  TOTAL_CARGA: 'Carga total',
  CONSUMO_kW: 'Consumo (kW)',
  CONSUMO_kW_MES: 'Consumo (kW/mês)',
  ATUALIZACAO: 'Atualização em campo',
  BLOQ_AMP: 'Bloq/Amp',
};

const MANAGER_FIELD_LABELS = {
  ...FIELD_LABELS,
  'ATUALIZAÃ‡': 'Atualização',
  DATA_ATUALIZACAO_DEIL: 'Atualização DEIL',
  TEM_IMAGEM: 'Tem imagem',
  LINKS_IMAGENS: 'Links imagens',
  LINK_IMAGEM_1: 'Imagem 1',
  LINK_IMAGEM_2: 'Imagem 2',
  LINK_IMAGEM_3: 'Imagem 3',
  LINK_IMAGEM_4: 'Imagem 4',
  LINK_IMAGEM_5: 'Imagem 5',
  POTENCIA_LUMINARIA_1: 'Pot. lum. 1',
  POTENCIA_LUMINARIA_2: 'Pot. lum. 2',
  POTENCIA_LUMINARIA_3: 'Pot. lum. 3',
  POTENCIA_LUMINARIA_4: 'Pot. lum. 4',
  POTENCIA_LUMINARIA_5: 'Pot. lum. 5',
  SYNCED_EM: 'Sincronizado em',
  EXPORTED_EM: 'Exportado em',
};

const MAX_IMAGE_SIZE_BYTES = 850 * 1024;
const MAX_LUMINARIAS = 5;

const FIELD_PLACEHOLDERS = {
  ID_CADASTRO: 'Opcional, se já existir no cadastro',
  TIPO_IMPLANTACAO: 'Selecione se é nova implantação ou poste existente',
  MOTIVO_IMPLANTACAO: 'Selecione o motivo',
  OBRA_NOME: 'Selecione ou informe a obra',
  IMPLANTACAO_CONCLUIDA: 'Marque se a implantação foi concluída',
  RPA: 'Preenchido automaticamente quando possível',
  LOCALIZACA: 'Selecione o tipo de localização',
  ENDERECO: 'Preenchido automaticamente pelas coordenadas',
  BARRAMENTO: 'Se houver',
  BAIRRO: 'Preenchido automaticamente quando possível',
  MEDICAO: 'Opcional: SIM ou NÃO',
  TIPO_LUMIN: 'Selecione o tipo de luminária',
  TIPO_DE_PO: 'Selecione o tipo de poste',
  TIPO_LAMPA: 'Selecione o tipo de lâmpada',
  QTDE: '1',
  BLOQ_AMP: 'Opcional',
};

const SESSION_STORAGE_KEY = 'luz-de-campo-session';
const MODULE_STORAGE_KEY = 'luz-de-campo-active-module';
const QUEUE_STORAGE_KEY = 'luz-de-campo-queue';
const SAO_JOAO_QUEUE_STORAGE_KEY = 'luz-de-campo-sao-joao-queue';
const LOCATION_CONTEXT_CACHE_KEY = 'luz-de-campo-location-context';
const OFFLINE_ACCESS_STORAGE_KEY = 'luz-de-campo-offline-access';
const OPERATORS_CACHE_STORAGE_KEY = 'luz-de-campo-operators-cache';
const PRODUCTION_FIELD_API_BASE_URL = 'https://cadastroeditor-api.onrender.com';

function resolveFieldApiBaseUrl(rawUrl) {
  const cleanedUrl = (rawUrl || 'http://127.0.0.1:8010').replace(/\/$/, '');

  if (typeof window === 'undefined') {
    return cleanedUrl;
  }

  const isProductionFrontend = window.location.hostname === 'luz-de-campo-frontend.onrender.com';
  const isLegacyBackendUrl = cleanedUrl === 'https://luz-de-campo.onrender.com';
  const isLocalBackendUrl = isLocalApiBaseUrl(cleanedUrl);

  if (isProductionFrontend && (isLegacyBackendUrl || isLocalBackendUrl)) {
    return PRODUCTION_FIELD_API_BASE_URL;
  }

  return cleanedUrl;
}

const FIELD_API_BASE_URL = resolveFieldApiBaseUrl(import.meta.env.VITE_FIELD_API_BASE_URL);
const MAX_OFFLINE_RPAS = 3;
const DEFAULT_FIELD_MODULES = [
  {
    slug: 'implantacao-obra',
    name: 'IMPLANTAÇÃO/OBRA',
    description: 'Formulário atual para registrar pontos de iluminação.',
  },
  {
    slug: 'sao-joao',
    name: 'SÃO JOÃO',
    description: 'Checklist temporário para polos e estruturas do evento.',
  },
];
const MAP_TILE_URLS = [
  'https://a.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}@2x.png',
  'https://b.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}@2x.png',
  'https://c.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}@2x.png',
  'https://d.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}@2x.png',
];

const QUICK_OPTIONS = {
  TIPO_IMPLANTACAO: ['NOVA IMPLANTAÇÃO', 'POSTE EXISTENTE'],
  MOTIVO_IMPLANTACAO: ['OBRA', 'EXPANSÃO'],
  OBRA_NOME: ['PEDONAL CAXANGÁ', 'PEDONAL ZONA OESTE'],
  IMPLANTACAO_CONCLUIDA: ['SIM', 'NÃO'],
  LOCALIZACA: ['RUA', 'PRAÇA', 'AVENIDA', 'OUTRO'],
  BARRAMENTO: ['S/N'],
  MEDICAO: ['SIM', 'NÃO'],
  TIPO_LUMIN: ['FECHADA', 'ORNAMENTAL', 'REFLETOR'],
  TIPO_DE_PO: ['CONCRETO', 'METÁLICO', 'FIBRA', 'GIRAFA'],
  TIPO_LAMPA: ['LED', 'V. SÓDIO', { label: 'METAL', value: 'V. METÁLICO' }, 'HALÓGENA'],
  QTDE: ['1', '2', '3', '4', '5'],
};

const FAST_COPY_FIELDS = [
  'TIPO_IMPLANTACAO',
  'MOTIVO_IMPLANTACAO',
  'OBRA_NOME',
  'RPA',
  'LOCALIZACA',
  'BAIRRO',
  'MEDICAO',
  'TIPO_LUMIN',
  'TIPO_DE_PO',
  'TIPO_LAMPA',
  'QTDE',
  'PERDAS',
  'TOTAL_CARGA',
  'CONSUMO_kW',
  'CONSUMO_kW_MES',
  'LUMINARIAS',
];

const FORM_FIELDS = [
  'TIPO_IMPLANTACAO',
  'MOTIVO_IMPLANTACAO',
  'OBRA_NOME',
  'IMPLANTACAO_CONCLUIDA',
  'RPA',
  'LOCALIZACA',
  'ENDERECO',
  'BARRAMENTO',
  'BAIRRO',
  'MEDICAO',
  'TIPO_LUMIN',
  'TIPO_DE_PO',
  'TIPO_LAMPA',
  'QTDE',
];

const SAO_JOAO_ANSWER_OPTIONS = ['SIM', 'NÃO', 'NÃO SE APLICA'];
const SAO_JOAO_QUESTIONS = [
  {
    group: 'Documentação técnica',
    items: [
      { key: 'q01_art', label: '1. A infraestrutura possui ART?', photos: [{ key: 'foto_art', label: 'Foto da ART' }] },
      { key: 'q02_projeto_art', label: '2. O projeto elétrico com os equipamentos foi encaminhado para a EMLURB junto com a ART?' },
      { key: 'q03_responsavel_presente', label: '3. O responsável técnico está presente durante a montagem?' },
      { key: 'q04_neoenergia', label: '4. O fornecimento da energia é feito pela Neoenergia PE?' },
    ],
  },
  {
    group: 'Proteções e cargas',
    items: [
      { key: 'q05_cargas_rede_ip', label: '5. Existem cargas conectadas diretamente à rede elétrica de IP?', photos: [{ key: 'foto_cargas_rede_ip', label: 'Foto das cargas conectadas à rede de IP' }] },
      { key: 'q06_disjuntor', label: '6. A instalação possui disjuntor termomagnético?', photos: [{ key: 'foto_sem_disjuntor', label: 'Foto do quadro sem disjuntor' }] },
      { key: 'q07_dps', label: '7. A instalação possui DPS?', photos: [{ key: 'foto_sem_dps', label: 'Foto do quadro sem DPS' }] },
      { key: 'q08_dr', label: '8. A instalação possui DR?', photos: [{ key: 'foto_sem_dr', label: 'Foto do quadro sem DR' }] },
      { key: 'q09_dr_30ma', label: '9. O DR instalado é de alta sensibilidade (30mA)?', photos: [{ key: 'foto_dr_baixa_sensibilidade', label: 'Foto do DR com baixa sensibilidade' }] },
    ],
  },
  {
    group: 'Aterramento',
    items: [
      { key: 'q10_partes_metalicas_aterradas', label: '10. As partes metálicas da estrutura estão aterradas?', photos: [{ key: 'foto_sem_aterramento', label: 'Foto do local sem aterramento' }] },
      {
        key: 'q11_aterramento_adequado',
        label: '11. O aterramento está adequado?',
        photos: [
          { key: 'foto_estrutura_aterrada', label: 'Foto da estrutura metálica aterrada' },
          { key: 'foto_aterramento_inadequado', label: 'Foto da estrutura metálica com aterramento inadequado' },
        ],
      },
      { key: 'q12_copperweld', label: '12. O aterramento usa vara Copperweld 5/8x2,40m, conector GAR/GTDU ou solda exotérmica e cabo verde mínimo 16mm²?', photos: [{ key: 'foto_aterramento', label: 'Foto do aterramento' }] },
      { key: 'q13_aterramento_massas', label: '13. O aterramento das massas foi respeitado?', photos: [{ key: 'foto_aterramento_massas', label: 'Foto do aterramento das massas' }] },
    ],
  },
  {
    group: 'Cabeamento e interferências',
    items: [
      { key: 'q14_cabos_expostos', label: '14. Existem cabeamentos expostos ao tempo, no chão ou amarrados em estruturas metálicas?', photos: [{ key: 'foto_cabos_expostos', label: 'Foto do cabeamento exposto, no chão ou amarrado' }] },
      { key: 'q15_trilho_passacabo', label: '15. Foi usado trilho passa-cabo protegendo os cabeamentos?', photos: [{ key: 'foto_cabos_sem_protecao', label: 'Foto dos cabos sem proteção' }] },
      { key: 'q16_iluminacao_provisoria', label: '16. Existem projetores, refletores ou iluminação provisória em postes metálicos exclusivos de IP ou árvores?', photos: [{ key: 'foto_iluminacao_provisoria', label: 'Foto da iluminação provisória em poste metálico ou árvore' }] },
      { key: 'q17_equipamentos_qualidade', label: '17. Os equipamentos usados nas instalações elétricas provisórias são de boa qualidade?' },
      { key: 'q18_elementos_ip', label: '18. Existem bandeiras, barracas, refletores, fiações ou outros elementos nos postes/equipamentos de IP?', photos: [{ key: 'foto_elementos_ip', label: 'Foto dos elementos no sistema de IP' }] },
    ],
  },
  {
    group: 'Riscos e violações',
    items: [
      { key: 'q19_distancias_at_mt', label: '19. As distâncias das estruturas para AT e MT estão respeitadas?', photos: [{ key: 'foto_distancias_at_mt', label: 'Foto das distâncias para rede AT ou MT' }] },
      { key: 'q20_caixas_violadas', label: '20. As caixas de passagem subterrâneas estão violadas?', photos: [{ key: 'foto_caixas_violadas', label: 'Foto das caixas violadas' }] },
      { key: 'q21_medidores_violados', label: '21. Os medidores da Neoenergia estão violados?', photos: [{ key: 'foto_medidores_violados', label: 'Foto dos medidores violados' }] },
      { key: 'q22_gambiarras', label: '22. Existem cordões luminosos ou gambiarras cruzando vias, grades ou elementos metálicos?', photos: [{ key: 'foto_elementos_cruzando', label: 'Foto dos elementos cruzando vias ou estruturas metálicas' }] },
    ],
  },
];

const SAO_JOAO_PHOTO_FIELDS = SAO_JOAO_QUESTIONS.flatMap((group) => group.items.flatMap((item) => item.photos || []));

function clampLuminaireCount(value) {
  const parsed = Number.parseInt(String(value ?? '').trim(), 10);
  if (!Number.isFinite(parsed)) return 1;
  return Math.min(MAX_LUMINARIAS, Math.max(1, parsed));
}

function createEmptyLuminaire(index) {
  return {
    INDICE: index,
    POTENCIA: '',
    IMAGEM: null,
  };
}

function normalizeLuminaireImage(image) {
  if (!image || typeof image !== 'object') return null;
  return image;
}

function normalizeLuminaireItems(rawItems, qtde, fallbackPotencia = '', fallbackImage = null) {
  const count = clampLuminaireCount(qtde || rawItems?.length || 1);
  const items = Array.isArray(rawItems) ? rawItems : [];

  return Array.from({ length: count }, (_, index) => {
    const rawItem = items[index] || {};
    const indice = index + 1;
    const potencia = rawItem.POTENCIA ?? rawItem.potencia ?? (indice === 1 ? fallbackPotencia : '') ?? '';
    const imagem = normalizeLuminaireImage(rawItem.IMAGEM ?? rawItem.imagem ?? (indice === 1 ? fallbackImage : null));

    return {
      INDICE: indice,
      POTENCIA: potencia === null || potencia === undefined ? '' : String(potencia),
      IMAGEM: imagem,
    };
  });
}

function countLuminaireImages(items) {
  return items.filter((item) => item?.IMAGEM).length;
}

function computeLuminairePotenciaTotal(items) {
  return items.reduce((total, item) => {
    const value = toNumber(item?.POTENCIA);
    return total + (Number.isFinite(value) ? value : 0);
  }, 0);
}

function normalizeEntryShape(entry) {
  const luminarias = normalizeLuminaireItems(entry?.LUMINARIAS ?? entry?.luminarias, entry?.QTDE ?? entry?.qtde, entry?.POTENCIA ?? entry?.potencia, entry?.IMAGEM ?? entry?.imagem);
  return {
    ...entry,
    QTDE: String(clampLuminaireCount(entry?.QTDE ?? entry?.qtde ?? luminarias.length)),
    LUMINARIAS: luminarias,
    IMAGEM: null,
  };
}

function formatCoordinate(value) {
  return Number(value).toFixed(6);
}

function formatAccuracy(value) {
  if (!Number.isFinite(value)) return '';
  return `${Math.round(value)} m`;
}

function parseManualCoordinates(rawValue) {
  const text = String(rawValue ?? '').trim();
  if (!text) return null;

  const matches = text.match(/-?\d+(?:[.,]\d+)?/g);
  if (!matches || matches.length < 2) return null;

  const lat = Number(matches[0].replace(',', '.'));
  const lng = Number(matches[1].replace(',', '.'));

  if (!Number.isFinite(lat) || !Number.isFinite(lng)) return null;
  if (Math.abs(lat) > 90 || Math.abs(lng) > 180) return null;

  return { lat, lng };
}

function toNumber(value) {
  if (value === null || value === undefined || value === '') return Number.NaN;
  return Number(String(value).replace(',', '.'));
}

function createEmptyFeatureCollection() {
  return { type: 'FeatureCollection', features: [] };
}

function buildPointCollection(lat, lng) {
  if (!Number.isFinite(lat) || !Number.isFinite(lng)) return createEmptyFeatureCollection();
  return {
    type: 'FeatureCollection',
    features: [{
      type: 'Feature',
      geometry: { type: 'Point', coordinates: [lng, lat] },
      properties: {},
    }],
  };
}

function exportWorkbook(rows) {
  const formattedRows = rows.map((row) => buildExportRow(row));
  const worksheet = XLSX.utils.json_to_sheet(
    formattedRows.map((row) => Object.fromEntries(EXPORT_FIELDS.map((field) => [field, row[field] ?? '']))),
    { header: EXPORT_FIELDS }
  );
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'LuzDeCampo');
  XLSX.writeFileXLSX(workbook, `luz_de_campo_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function buildExportRow(row) {
  const luminarias = normalizeLuminaireItems(
    row?.LUMINARIAS ?? row?.luminarias ?? row?.luminarias_json,
    row?.QTDE ?? row?.qtde,
    row?.POTENCIA ?? row?.potencia,
    row?.IMAGEM ?? row?.imagem
  );
  const qtde = String(clampLuminaireCount(row?.QTDE ?? row?.qtde ?? luminarias.length ?? 1));
  const potenciaTotal = row?.POTENCIA ?? row?.potencia ?? computeLuminairePotenciaTotal(luminarias);
  const imageLinksByLuminaire = row?.image_links_by_luminaria || {};
  const imageLinks = Array.isArray(row?.image_links)
    ? row.image_links
    : String(row?.LINKS_IMAGENS ?? '')
        .split('\n')
        .map((item) => item.trim())
        .filter(Boolean);
  const exportRow = {
    RPA: row?.RPA ?? row?.rpa ?? '',
    LOCALIZACA: row?.LOCALIZACA ?? row?.localizacao ?? '',
    ENDERECO: row?.ENDERECO ?? row?.endereco ?? '',
    BARRAMENTO: row?.BARRAMENTO ?? row?.barramento ?? '',
    LATITUDE: row?.LATITUDE ?? row?.latitude ?? '',
    LONGITUDE: row?.LONGITUDE ?? row?.longitude ?? '',
    BAIRRO: row?.BAIRRO ?? row?.bairro ?? '',
    MEDICAO: row?.MEDICAO ?? row?.medicao ?? '',
    TIPO_LUMIN: row?.TIPO_LUMIN ?? row?.tipo_lumin ?? '',
    TIPO_DE_PO: row?.TIPO_DE_PO ?? row?.tipo_de_po ?? '',
    TIPO_LAMPA: row?.TIPO_LAMPA ?? row?.tipo_lampa ?? '',
    QTDE: qtde,
    POTENCIA: potenciaTotal ?? '',
    PERDAS: row?.PERDAS ?? row?.perdas ?? '',
    TOTAL_CARGA: row?.TOTAL_CARGA ?? row?.total_carga ?? '',
    CONSUMO_kW: row?.CONSUMO_kW ?? row?.consumo_kw ?? '',
    CONSUMO_kW_MES: row?.CONSUMO_kW_MES ?? row?.consumo_kw_mes ?? '',
    'ATUALIZAÇ': row?.['ATUALIZAÇ'] ?? row?.ATUALIZACAO ?? row?.atualizacao ?? '',
    DATA_ATUALIZACAO_DEIL: row?.DATA_ATUALIZACAO_DEIL ?? '',
    BLOQ_AMP: row?.BLOQ_AMP ?? row?.bloq_amp ?? '',
    OPERADOR: row?.OPERADOR ?? row?.operador ?? '',
    ID_CADASTRO: row?.ID_CADASTRO ?? row?.id_cadastro ?? '',
    TIPO_IMPLANTACAO: row?.TIPO_IMPLANTACAO ?? row?.tipo_implantacao ?? '',
    MOTIVO_IMPLANTACAO: row?.MOTIVO_IMPLANTACAO ?? row?.motivo_implantacao ?? '',
    OBRA_NOME: row?.OBRA_NOME ?? row?.obra_nome ?? '',
    IMPLANTACAO_CONCLUIDA: row?.IMPLANTACAO_CONCLUIDA ?? row?.implantacao_concluida ?? '',
    TEM_IMAGEM: row?.TEM_IMAGEM ?? (row?.tem_imagem ? 'SIM' : countLuminaireImages(luminarias) > 0 ? 'SIM' : 'NÃO'),
    LINKS_IMAGENS: row?.LINKS_IMAGENS ?? (imageLinks.length ? imageLinks.join('\n') : ''),
    SYNCED_EM: row?.SYNCED_EM ?? row?.synced_em ?? '',
    EXPORTED_EM: row?.EXPORTED_EM ?? row?.exported_em ?? '',
  };

  for (let index = 0; index < MAX_LUMINARIAS; index += 1) {
    const item = luminarias[index];
    const columnIndex = index + 1;
    exportRow[`POTENCIA_LUMINARIA_${columnIndex}`] = item?.POTENCIA ?? item?.potencia ?? '';
    exportRow[`LINK_IMAGEM_${columnIndex}`] =
      row?.[`LINK_IMAGEM_${columnIndex}`]
      ?? imageLinksByLuminaire[String(columnIndex)]
      ?? '';
  }

  return exportRow;
}

function formatManagerCellValue(value) {
  if (value === null || value === undefined || value === '') return '-';
  if (typeof value === 'number') return String(value);
  const text = String(value).trim();
  if (!text) return '-';

  if (/^\d{4}-\d{2}-\d{2}T/.test(text)) {
    const date = new Date(text);
    if (!Number.isNaN(date.getTime())) {
      return date.toLocaleString('pt-BR');
    }
  }

  return text;
}

function isManagerImageLinkField(field) {
  return field === 'LINKS_IMAGENS' || field.startsWith('LINK_IMAGEM_');
}

function isEntryReadyToSync(entry) {
  return (
    Number.isFinite(toNumber(entry?.LATITUDE))
    && Number.isFinite(toNumber(entry?.LONGITUDE))
    && entry?.IMPLANTACAO_CONCLUIDA === 'SIM'
  );
}

function buildFormFocusSequence(currentForm) {
  const baseFields = FORM_FIELDS.filter((field) => (
    !['QTDE', 'LATITUDE', 'LONGITUDE'].includes(field)
    && (field !== 'OBRA_NOME' || currentForm?.MOTIVO_IMPLANTACAO === 'OBRA')
  ));
  const luminariaFields = (currentForm?.LUMINARIAS || []).flatMap((item, index) => [
    `LUMINARIA_${item?.INDICE || index + 1}_POTENCIA`,
    `LUMINARIA_${item?.INDICE || index + 1}_IMAGEM`,
  ]);
  return [...baseFields, ...luminariaFields];
}

function isFocusKeyFilled(currentForm, focusKey) {
  if (focusKey.startsWith('LUMINARIA_')) {
    const match = focusKey.match(/^LUMINARIA_(\d+)_(POTENCIA|IMAGEM)$/);
    if (!match) return false;
    const luminariaIndex = Number(match[1]) - 1;
    const luminaria = currentForm?.LUMINARIAS?.[luminariaIndex];
    if (!luminaria) return false;
    if (match[2] === 'POTENCIA') {
      return String(luminaria.POTENCIA ?? '').trim().length > 0;
    }
    return Boolean(luminaria.IMAGEM);
  }

  return String(currentForm?.[focusKey] ?? '').trim().length > 0;
}

function buildNextPointForm(baseForm, operatorName) {
  const luminarias = normalizeLuminaireItems(
    baseForm?.LUMINARIAS ?? baseForm?.luminarias,
    baseForm?.QTDE ?? baseForm?.qtde ?? 1,
    baseForm?.POTENCIA ?? baseForm?.potencia,
    baseForm?.IMAGEM ?? baseForm?.imagem
  );
  return {
    ...INITIAL_FORM,
    ...baseForm,
    OPERADOR: operatorName || baseForm?.OPERADOR || '',
    ID_CADASTRO: '',
    ENDERECO: '',
    LATITUDE: '',
    LONGITUDE: '',
    BAIRRO: '',
    RPA: '',
    QTDE: String(clampLuminaireCount(baseForm?.QTDE ?? baseForm?.qtde ?? luminarias.length)),
    LUMINARIAS: luminarias.map((item, index) => ({
      INDICE: index + 1,
      POTENCIA: item.POTENCIA ?? item.potencia ?? '',
      IMAGEM: null,
    })),
    IMAGEM: null,
    OBRA_NOME: baseForm?.MOTIVO_IMPLANTACAO === 'OBRA' ? baseForm?.OBRA_NOME || '' : '',
    IMPLANTACAO_CONCLUIDA: '',
    ATUALIZACAO: new Date().toISOString().slice(0, 10),
  };
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ''));
    reader.onerror = () => reject(new Error('Não foi possível ler a imagem.'));
    reader.readAsDataURL(blob);
  });
}

function loadImageElementFromUrl(objectUrl) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error('Não foi possível processar a imagem.'));
    image.src = objectUrl;
  });
}

function canvasToJpegBlob(canvas, quality) {
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => {
        if (!blob) {
          reject(new Error('Não foi possível compactar a imagem.'));
          return;
        }
        resolve(blob);
      },
      'image/jpeg',
      quality
    );
  });
}

function getImageProcessingPreset(file) {
  const deviceMemory = typeof navigator !== 'undefined' && Number.isFinite(Number(navigator.deviceMemory))
    ? Number(navigator.deviceMemory)
    : null;
  const fileSizeMb = file?.size ? file.size / (1024 * 1024) : 0;

  if ((deviceMemory !== null && deviceMemory <= 2) || fileSizeMb >= 8) {
    return { maxDimension: 960, qualities: [0.72, 0.62, 0.52, 0.44] };
  }

  if ((deviceMemory !== null && deviceMemory <= 4) || fileSizeMb >= 4) {
    return { maxDimension: 1120, qualities: [0.78, 0.68, 0.58, 0.5] };
  }

  return { maxDimension: 1280, qualities: [0.82, 0.74, 0.66, 0.58] };
}

function releaseCanvas(canvas) {
  canvas.width = 1;
  canvas.height = 1;
}

async function compressImageWithPreset(sourceImage, preset) {
  const canvas = document.createElement('canvas');
  const ratio = Math.min(1, preset.maxDimension / Math.max(sourceImage.width, sourceImage.height));
  canvas.width = Math.max(1, Math.round(sourceImage.width * ratio));
  canvas.height = Math.max(1, Math.round(sourceImage.height * ratio));
  const context = canvas.getContext('2d');
  if (!context) {
    releaseCanvas(canvas);
    throw new Error('Não foi possível preparar a imagem.');
  }

  context.drawImage(sourceImage, 0, 0, canvas.width, canvas.height);

  let compressedBlob = await canvasToJpegBlob(canvas, preset.qualities[0]);
  for (const quality of preset.qualities.slice(1)) {
    if (compressedBlob.size <= MAX_IMAGE_SIZE_BYTES) break;
    compressedBlob = await canvasToJpegBlob(canvas, quality);
  }

  if (compressedBlob.size > MAX_IMAGE_SIZE_BYTES) {
    releaseCanvas(canvas);
    throw new Error('A imagem ficou muito pesada. Tente uma foto mais leve.');
  }

  const compressedDataUrl = await blobToDataUrl(compressedBlob);
  releaseCanvas(canvas);

  return {
    data_url: compressedDataUrl,
    size_bytes: compressedBlob.size,
  };
}

async function prepareImagePayload(file) {
  const objectUrl = URL.createObjectURL(file);

  try {
    const sourceImage = await loadImageElementFromUrl(objectUrl);
    const primaryPreset = getImageProcessingPreset(file);

    try {
      const compressed = await compressImageWithPreset(sourceImage, primaryPreset);
      return {
        client_id: crypto.randomUUID(),
        filename: (file.name || 'foto.jpg').replace(/\.[^.]+$/, '.jpg'),
        content_type: 'image/jpeg',
        ...compressed,
      };
    } catch {
      const fallbackPreset = {
        maxDimension: Math.max(640, Math.round(primaryPreset.maxDimension * 0.75)),
        qualities: [0.62, 0.54, 0.48, 0.42],
      };
      const compressed = await compressImageWithPreset(sourceImage, fallbackPreset);
      return {
        client_id: crypto.randomUUID(),
        filename: (file.name || 'foto.jpg').replace(/\.[^.]+$/, '.jpg'),
        content_type: 'image/jpeg',
        ...compressed,
      };
    }
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

function getStoredOperator() {
  if (typeof window === 'undefined') return null;

  try {
    const rawSession = window.localStorage.getItem(SESSION_STORAGE_KEY);
    if (!rawSession) return null;

    return JSON.parse(rawSession);
  } catch {
    return null;
  }
}

function getStoredModule() {
  if (typeof window === 'undefined') return null;

  try {
    return window.localStorage.getItem(MODULE_STORAGE_KEY) || null;
  } catch {
    return null;
  }
}

function getStoredQueue() {
  if (typeof window === 'undefined') return [];

  try {
    const rawEntries = window.localStorage.getItem(QUEUE_STORAGE_KEY);
    if (!rawEntries) return [];
    const parsed = JSON.parse(rawEntries);
    return Array.isArray(parsed) ? parsed.map((entry) => normalizeEntryShape(entry)) : [];
  } catch {
    return [];
  }
}

function getStoredSaoJoaoQueue() {
  if (typeof window === 'undefined') return [];

  try {
    const rawEntries = window.localStorage.getItem(SAO_JOAO_QUEUE_STORAGE_KEY);
    if (!rawEntries) return [];
    const parsed = JSON.parse(rawEntries);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function createInitialSaoJoaoForm(operator) {
  return {
    POLO_NOME: '',
    LATITUDE: '',
    LONGITUDE: '',
    OBSERVACOES: '',
    RESPOSTAS: {},
    IMAGENS: {},
    OPERADOR: operator?.name || '',
    OPERADOR_ID: operator?.id || null,
  };
}

function buildSaoJoaoPayload(entry) {
  const images = SAO_JOAO_PHOTO_FIELDS
    .map((photoField) => {
      const image = entry.IMAGENS?.[photoField.key];
      if (!image) return null;
      return {
        ...image,
        field_key: photoField.key,
        field_label: photoField.label,
      };
    })
    .filter(Boolean);

  return {
    CLIENT_UUID: entry.CLIENT_UUID,
    OPERADOR_ID: entry.OPERADOR_ID,
    OPERADOR: entry.OPERADOR,
    POLO_NOME: entry.POLO_NOME,
    LATITUDE: entry.LATITUDE,
    LONGITUDE: entry.LONGITUDE,
    RESPOSTAS: entry.RESPOSTAS || {},
    OBSERVACOES: entry.OBSERVACOES,
    IMAGENS: images,
  };
}

function roundCoordinate(value) {
  return Number.isFinite(value) ? Number(value).toFixed(5) : '';
}

function buildLocationContextCacheKey(lat, lng) {
  return `${roundCoordinate(lat)}:${roundCoordinate(lng)}`;
}

function getStoredLocationContextCache() {
  if (typeof window === 'undefined') return {};

  try {
    const rawCache = window.localStorage.getItem(LOCATION_CONTEXT_CACHE_KEY);
    if (!rawCache) return {};
    const parsed = JSON.parse(rawCache);
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch {
    return {};
  }
}

function storeLocationContext(lat, lng, context) {
  if (typeof window === 'undefined') return;

  try {
    const nextCache = {
      ...getStoredLocationContextCache(),
      [buildLocationContextCacheKey(lat, lng)]: context,
    };
    window.localStorage.setItem(LOCATION_CONTEXT_CACHE_KEY, JSON.stringify(nextCache));
  } catch {
    // Cache is best-effort only.
  }
}

function getStoredLocationContext(lat, lng) {
  const cache = getStoredLocationContextCache();
  return cache[buildLocationContextCacheKey(lat, lng)] || null;
}

function getStoredAllowedOperatorsCache() {
  if (typeof window === 'undefined') return [];

  try {
    const rawValue = window.localStorage.getItem(OPERATORS_CACHE_STORAGE_KEY);
    if (!rawValue) return [];
    const parsed = JSON.parse(rawValue);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function storeAllowedOperatorsCache(operators) {
  if (typeof window === 'undefined') return;

  try {
    window.localStorage.setItem(OPERATORS_CACHE_STORAGE_KEY, JSON.stringify(operators));
  } catch {
    // Cache is best-effort only.
  }
}

function getStoredOfflineAccessProfiles() {
  if (typeof window === 'undefined') return [];

  try {
    const rawValue = window.localStorage.getItem(OFFLINE_ACCESS_STORAGE_KEY);
    if (!rawValue) return [];
    const parsed = JSON.parse(rawValue);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function isLocalDevelopmentHost(hostname) {
  return hostname === 'localhost' || hostname === '127.0.0.1';
}

function isLocalApiBaseUrl(url) {
  return /https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/i.test(url);
}

async function readApiError(response, fallbackMessage) {
  try {
    const contentType = response.headers.get('content-type') || '';
    if (contentType.includes('application/json')) {
      const payload = await response.json();
      if (typeof payload?.detail === 'string' && payload.detail.trim()) {
        return payload.detail.trim();
      }
      if (typeof payload?.message === 'string' && payload.message.trim()) {
        return payload.message.trim();
      }
    } else {
      const text = await response.text();
      if (text.trim()) return text.trim();
    }
  } catch {
    // Fallback below.
  }

  return fallbackMessage;
}

function storeOfflineAccessProfile(profile) {
  if (typeof window === 'undefined') return;

  try {
    const currentProfiles = getStoredOfflineAccessProfiles().filter((item) => item?.id !== profile.id);
    currentProfiles.push(profile);
    window.localStorage.setItem(OFFLINE_ACCESS_STORAGE_KEY, JSON.stringify(currentProfiles));
  } catch {
    // Cache is best-effort only.
  }
}

function buildMergedOperatorList(onlineOperators = []) {
  const merged = new Map();

  [...getStoredAllowedOperatorsCache(), ...onlineOperators].forEach((operator) => {
    if (!operator?.id) return;
    merged.set(operator.id, operator);
  });

  getStoredOfflineAccessProfiles().forEach((profile) => {
    if (!profile?.id) return;
    const current = merged.get(profile.id) || {};
    merged.set(profile.id, {
      ...current,
      id: profile.id,
      name: profile.name || current.name || `Operador ${profile.id}`,
      can_export: Boolean(profile.can_export ?? current.can_export),
      offline_ready: true,
      offline_rpa: Array.isArray(profile.prepared_rpas)
        ? profile.prepared_rpas.join(', ')
        : (profile.prepared_rpa || null),
    });
  });

  return Array.from(merged.values()).sort((first, second) => String(first.name || '').localeCompare(String(second.name || ''), 'pt-BR'));
}

async function hashAccessCode(accessCode) {
  const buffer = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(accessCode));
  return Array.from(new Uint8Array(buffer), (value) => value.toString(16).padStart(2, '0')).join('');
}

function tileXForLng(lng, zoom) {
  return Math.floor(((lng + 180) / 360) * (2 ** zoom));
}

function tileYForLat(lat, zoom) {
  const radians = (lat * Math.PI) / 180;
  const value = (1 - Math.log(Math.tan(radians) + 1 / Math.cos(radians)) / Math.PI) / 2;
  return Math.floor(value * (2 ** zoom));
}

function clampTileIndex(value, zoom) {
  const maxIndex = (2 ** zoom) - 1;
  return Math.max(0, Math.min(maxIndex, value));
}

function collectGeometryCoordinates(geometry, sink = []) {
  if (!geometry) return sink;

  if (typeof geometry[0] === 'number' && typeof geometry[1] === 'number') {
    sink.push(geometry);
    return sink;
  }

  if (Array.isArray(geometry)) {
    geometry.forEach((item) => collectGeometryCoordinates(item, sink));
  }

  return sink;
}

function getFeatureCollectionBounds(featureCollection) {
  const coordinates = (featureCollection?.features || []).flatMap((feature) => collectGeometryCoordinates(feature?.geometry?.coordinates, []));
  if (!coordinates.length) return null;

  let minLng = Number.POSITIVE_INFINITY;
  let minLat = Number.POSITIVE_INFINITY;
  let maxLng = Number.NEGATIVE_INFINITY;
  let maxLat = Number.NEGATIVE_INFINITY;

  coordinates.forEach(([lng, lat]) => {
    if (!Number.isFinite(lng) || !Number.isFinite(lat)) return;
    minLng = Math.min(minLng, lng);
    minLat = Math.min(minLat, lat);
    maxLng = Math.max(maxLng, lng);
    maxLat = Math.max(maxLat, lat);
  });

  if (![minLng, minLat, maxLng, maxLat].every(Number.isFinite)) return null;
  return { minLng, minLat, maxLng, maxLat };
}

function buildTileUrlsForBounds(bounds) {
  if (!bounds) return [];

  const zoomLevels = [12, 13, 14, 15, 16];
  const maxTiles = 900;
  const urls = [];

  for (const zoom of zoomLevels) {
    const minX = clampTileIndex(tileXForLng(bounds.minLng, zoom) - 1, zoom);
    const maxX = clampTileIndex(tileXForLng(bounds.maxLng, zoom) + 1, zoom);
    const minY = clampTileIndex(tileYForLat(bounds.maxLat, zoom) - 1, zoom);
    const maxY = clampTileIndex(tileYForLat(bounds.minLat, zoom) + 1, zoom);

    for (let x = minX; x <= maxX; x += 1) {
      for (let y = minY; y <= maxY; y += 1) {
        if (urls.length >= maxTiles) return urls;
        const template = MAP_TILE_URLS[urls.length % MAP_TILE_URLS.length];
        urls.push(
          template
            .replace('{z}', String(zoom))
            .replace('{x}', String(x))
            .replace('{y}', String(y))
        );
      }
    }
  }

  return urls;
}

async function prefetchTilesForBounds(bounds) {
  const urls = buildTileUrlsForBounds(bounds);
  if (!urls.length) return 0;

  const chunkSize = 12;
  let loadedCount = 0;

  for (let index = 0; index < urls.length; index += chunkSize) {
    const batch = urls.slice(index, index + chunkSize);
    const responses = await Promise.allSettled(
      batch.map((url) => fetch(url, { mode: 'no-cors', cache: 'reload' }))
    );
    loadedCount += responses.filter((response) => response.status === 'fulfilled').length;
  }

  return loadedCount;
}

function mergeLocationContext(...contexts) {
  return contexts.reduce((accumulator, context) => {
    if (!context) return accumulator;
    return {
      endereco: accumulator.endereco || context.endereco || '',
      bairro: accumulator.bairro || context.bairro || '',
      rpa: accumulator.rpa || context.rpa || '',
    };
  }, { endereco: '', bairro: '', rpa: '' });
}

function extractReverseGeocodeAddress(payload) {
  const address = payload?.address || {};
  const street = address.road || address.pedestrian || address.footway || address.path || address.cycleway || '';
  const houseNumber = address.house_number || '';
  const label = [street, houseNumber].filter(Boolean).join(', ');
  return label || payload?.name || payload?.display_name?.split(',').slice(0, 2).join(', ') || '';
}

function extractReverseGeocodeDistrict(payload) {
  const address = payload?.address || {};
  return address.suburb
    || address.neighbourhood
    || address.city_district
    || address.quarter
    || address.borough
    || '';
}

async function fetchAllowedOperators() {
  const response = await fetch(`${FIELD_API_BASE_URL}/field-access/operators`);
  if (!response.ok) {
    throw new Error('Não foi possível carregar os usuários permitidos.');
  }
  const operators = await response.json();
  storeAllowedOperatorsCache(operators);
  return operators;
}

async function loginFieldOperator(operatorId, accessCode) {
  const response = await fetch(`${FIELD_API_BASE_URL}/field-access/login`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      operator_id: Number(operatorId),
      access_code: accessCode,
    }),
  });

  if (!response.ok) {
    throw new Error('Acesso não autorizado. Confira usuário e código.');
  }

  return response.json();
}

async function loginOfflineOperator(operatorId, accessCode) {
  const profile = getStoredOfflineAccessProfiles().find((item) => String(item?.id) === String(operatorId));
  if (!profile) {
    throw new Error('Esse acesso offline ainda não foi preparado neste aparelho.');
  }

  const accessHash = await hashAccessCode(accessCode);
  if (accessHash !== profile.access_hash) {
    throw new Error('Acesso não autorizado. Confira usuário e código.');
  }

  return {
    id: profile.id,
    name: profile.name,
    can_export: Boolean(profile.can_export),
  };
}

async function fetchRpaOptions() {
  const response = await fetch(`${FIELD_API_BASE_URL}/filter-options/rpa`);
  if (!response.ok) {
    throw new Error('Não foi possível carregar as RPAs.');
  }
  return response.json();
}

async function fetchRpaBoundary(rpas) {
  const params = new URLSearchParams();
  rpas.forEach((rpa) => params.append('rpa', rpa));
  const response = await fetch(`${FIELD_API_BASE_URL}/rpa-boundaries?${params.toString()}`);
  if (!response.ok) {
    throw new Error('Não foi possível preparar a área offline.');
  }
  return response.json();
}

async function prepareOfflineAccess(operator, accessCode, selectedRpas) {
  const boundary = await fetchRpaBoundary(selectedRpas);
  const bounds = getFeatureCollectionBounds(boundary);
  if (!bounds) {
    throw new Error('Não foi possível montar a área da RPA para uso offline.');
  }

  const accessHash = await hashAccessCode(accessCode);
  const cachedTiles = await prefetchTilesForBounds(bounds);

  storeOfflineAccessProfile({
    id: operator.id,
    name: operator.name,
    can_export: operator.can_export,
    access_hash: accessHash,
    prepared_rpas: selectedRpas,
    prepared_at: new Date().toISOString(),
    cached_tiles: cachedTiles,
    bounds,
  });

  return cachedTiles;
}

async function createFieldOperator(managerOperatorId, managerAccessCode, payload) {
  const response = await fetch(`${FIELD_API_BASE_URL}/field-access/users`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      manager_operator_id: managerOperatorId,
      manager_access_code: managerAccessCode,
      name: payload.name,
      access_code: payload.accessCode,
      can_export: payload.canExport,
    }),
  });

  if (!response.ok) {
    const errorPayload = await response.json().catch(() => null);
    throw new Error(errorPayload?.detail || 'Não foi possível cadastrar o usuário.');
  }

  return response.json();
}

async function deactivateFieldOperator(managerOperatorId, managerAccessCode, operatorId) {
  const response = await fetch(`${FIELD_API_BASE_URL}/field-access/users/${operatorId}/deactivate`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      manager_operator_id: managerOperatorId,
      manager_access_code: managerAccessCode,
    }),
  });

  if (!response.ok) {
    const errorPayload = await response.json().catch(() => null);
    throw new Error(errorPayload?.detail || 'Não foi possível desativar o usuário.');
  }

  return response.json();
}

async function fetchFieldModules({ includeInactive = false } = {}) {
  try {
    const params = includeInactive ? '?include_inactive=true' : '';
    const response = await fetch(`${FIELD_API_BASE_URL}/field-modules${params}`);
    if (!response.ok) {
      throw new Error('Não foi possível carregar os módulos.');
    }
    const modules = await response.json();
    return Array.isArray(modules) && modules.length ? modules : DEFAULT_FIELD_MODULES;
  } catch {
    return DEFAULT_FIELD_MODULES;
  }
}

async function updateFieldModule(moduleSlug, payload) {
  const response = await fetch(`${FIELD_API_BASE_URL}/field-modules/${encodeURIComponent(moduleSlug)}`, {
    method: 'PUT',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    throw new Error(await readApiError(response, 'Não foi possível atualizar o módulo.'));
  }

  return response.json();
}

async function syncFieldEntries(entries) {
  if (
    typeof window !== 'undefined'
    && !isLocalDevelopmentHost(window.location.hostname)
    && isLocalApiBaseUrl(FIELD_API_BASE_URL)
  ) {
    throw new Error('A API do Luz de Campo não foi configurada neste deploy. Defina VITE_FIELD_API_BASE_URL com a URL pública do backend.');
  }

  let response;
  try {
    response = await fetch(`${FIELD_API_BASE_URL}/field-submissions/sync`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ entries }),
    });
  } catch {
    throw new Error(`Não foi possível alcançar a API em ${FIELD_API_BASE_URL}. Verifique a URL pública do backend, CORS e disponibilidade do serviço.`);
  }

  if (!response.ok) {
    throw new Error(await readApiError(response, 'Não foi possível sincronizar os apontamentos.'));
  }

  return response.json();
}

async function syncSaoJoaoEntries(entries) {
  let response;
  try {
    response = await fetch(`${FIELD_API_BASE_URL}/sao-joao-submissions/sync`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ entries: entries.map((entry) => buildSaoJoaoPayload(entry)) }),
    });
  } catch {
    throw new Error(`Não foi possível alcançar a API em ${FIELD_API_BASE_URL}. Verifique o backend e tente novamente.`);
  }

  if (!response.ok) {
    throw new Error(await readApiError(response, 'Não foi possível sincronizar o checklist do São João.'));
  }

  return response.json();
}

async function fetchLocationContext(lat, lng) {
  const params = new URLSearchParams({
    lat: String(lat),
    lng: String(lng),
  });
  const cachedContext = getStoredLocationContext(lat, lng);

  let backendContext = null;
  try {
    const response = await fetch(`${FIELD_API_BASE_URL}/field-context?${params.toString()}`);
    if (response.ok) {
      backendContext = await response.json();
    }
  } catch {
    backendContext = null;
  }

  let reverseContext = null;
  const needsReverseLookup = !backendContext?.endereco || !backendContext?.bairro;
  if (needsReverseLookup && typeof navigator !== 'undefined' && navigator.onLine) {
    try {
      const reverseParams = new URLSearchParams({
        format: 'jsonv2',
        lat: String(lat),
        lon: String(lng),
        zoom: '18',
        addressdetails: '1',
        'accept-language': 'pt-BR',
      });
      const reverseResponse = await fetch(`https://nominatim.openstreetmap.org/reverse?${reverseParams.toString()}`);
      if (reverseResponse.ok) {
        const payload = await reverseResponse.json();
        reverseContext = {
          endereco: extractReverseGeocodeAddress(payload),
          bairro: extractReverseGeocodeDistrict(payload),
          rpa: '',
        };
      }
    } catch {
      reverseContext = null;
    }
  }

  const mergedContext = mergeLocationContext(backendContext, reverseContext, cachedContext);
  if (!mergedContext.endereco && !mergedContext.bairro && !mergedContext.rpa) {
    throw new Error('Não foi possível completar o endereço automaticamente.');
  }

  storeLocationContext(lat, lng, mergedContext);
  return mergedContext;
}

async function fetchSubmissionSummary() {
  const response = await fetch(`${FIELD_API_BASE_URL}/field-submissions/summary`);
  if (!response.ok) {
    throw new Error('Não foi possível carregar o resumo gerencial.');
  }
  return response.json();
}

async function fetchSubmissionExport(scope) {
  const response = await fetch(`${FIELD_API_BASE_URL}/field-submissions/export?scope=${scope}`);
  if (!response.ok) {
    throw new Error('Não foi possível carregar os dados para exportação.');
  }
  return response.json();
}

async function markSubmissionsExported(clientUuids, exportedByOperatorId) {
  const response = await fetch(`${FIELD_API_BASE_URL}/field-submissions/mark-exported`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      client_uuids: clientUuids,
      exported_by_operator_id: exportedByOperatorId,
    }),
  });

  if (!response.ok) {
    throw new Error('Não foi possível marcar os registros como exportados.');
  }

  return response.json();
}

async function updateFieldSubmission(clientUuid, payload) {
  const response = await fetch(`${FIELD_API_BASE_URL}/field-submissions/${encodeURIComponent(clientUuid)}`, {
    method: 'PUT',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    throw new Error(await readApiError(response, 'Não foi possível editar o ponto.'));
  }

  return response.json();
}

async function deleteFieldSubmission(clientUuid, payload) {
  const normalizedClientUuid = String(clientUuid || '').trim();
  if (!normalizedClientUuid) {
    throw new Error('Não foi possível identificar este ponto para exclusão.');
  }

  const deletePayload = {
    ...payload,
    client_uuid: normalizedClientUuid,
  };

  let response = await fetch(`${FIELD_API_BASE_URL}/field-submissions/delete`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(deletePayload),
  });

  if (response.status === 404) {
    response = await fetch(`${FIELD_API_BASE_URL}/field-submissions/${encodeURIComponent(normalizedClientUuid)}/delete`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(payload),
    });
  }

  if (response.status === 404) {
    response = await fetch(`${FIELD_API_BASE_URL}/field-submissions/delete/${encodeURIComponent(normalizedClientUuid)}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(payload),
    });
  }

  if (response.status === 404) {
    response = await fetch(`${FIELD_API_BASE_URL}/field-submissions/${encodeURIComponent(normalizedClientUuid)}`, {
      method: 'DELETE',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(payload),
    });
  }

  if (!response.ok) {
    throw new Error(await readApiError(response, 'Não foi possível excluir o ponto.'));
  }

  return response.json();
}

export default function App() {
  const mapContainerRef = useRef(null);
  const formSectionRef = useRef(null);
  const formFieldRefs = useRef({});
  const luminaireFileInputRefs = useRef({});
  const saveActionRef = useRef(null);
  const formStateRef = useRef(INITIAL_FORM);
  const handledAutofocusRequestRef = useRef(0);
  const mapRef = useRef(null);
  const markerRef = useRef(null);
  const [step, setStep] = useState('location');
  const [geoStatus, setGeoStatus] = useState('Pronto para captar a localização.');
  const [captureError, setCaptureError] = useState('');
  const [draftPosition, setDraftPosition] = useState(null);
  const [confirmedPosition, setConfirmedPosition] = useState(null);
  const [gpsAccuracy, setGpsAccuracy] = useState(null);
  const [manualCoordinateValue, setManualCoordinateValue] = useState('');
  const [manualCoordinateError, setManualCoordinateError] = useState('');
  const [form, setForm] = useState(INITIAL_FORM);
  const [entries, setEntries] = useState(() => getStoredQueue());
  const [editingEntryId, setEditingEntryId] = useState(null);
  const [toast, setToast] = useState('');
  const [activeOperator, setActiveOperator] = useState(() => getStoredOperator());
  const [activeModule, setActiveModule] = useState(() => getStoredModule());
  const [fieldModules, setFieldModules] = useState(DEFAULT_FIELD_MODULES);
  const [accessForm, setAccessForm] = useState({ operatorId: '', accessCode: '', offlineRpas: [] });
  const [authError, setAuthError] = useState('');
  const [authInfo, setAuthInfo] = useState('');
  const [allowedOperators, setAllowedOperators] = useState([]);
  const [authLoading, setAuthLoading] = useState(true);
  const [rpaOptions, setRpaOptions] = useState([]);
  const [offlinePrepLoading, setOfflinePrepLoading] = useState(false);
  const [syncing, setSyncing] = useState(false);
  const [queueFilter, setQueueFilter] = useState('all');
  const [nextPointReady, setNextPointReady] = useState(false);
  const [online, setOnline] = useState(() => (typeof navigator === 'undefined' ? true : navigator.onLine));
  const [managerSummary, setManagerSummary] = useState({ total: 0, pending_export: 0, exported: 0 });
  const [managerLoading, setManagerLoading] = useState(false);
  const [managerRows, setManagerRows] = useState([]);
  const [managerRowsLoading, setManagerRowsLoading] = useState(false);
  const [managerEditingClientUuid, setManagerEditingClientUuid] = useState('');
  const [managerEditDraft, setManagerEditDraft] = useState({});
  const [managerTableFilters, setManagerTableFilters] = useState({});
  const [managerTableMaximized, setManagerTableMaximized] = useState(false);
  const [managerUsersCollapsed, setManagerUsersCollapsed] = useState(true);
  const [managerRowActionId, setManagerRowActionId] = useState('');
  const [managerModuleActionSlug, setManagerModuleActionSlug] = useState('');
  const [managerUserForm, setManagerUserForm] = useState({
    managerAccessCode: '',
    name: '',
    accessCode: '',
    canExport: false,
  });
  const [managerUserLoading, setManagerUserLoading] = useState(false);
  const [managerUserActionId, setManagerUserActionId] = useState(null);
  const [autofocusRequestId, setAutofocusRequestId] = useState(0);
  const [guidedFieldKey, setGuidedFieldKey] = useState('');
  const [saoJoaoForm, setSaoJoaoForm] = useState(() => createInitialSaoJoaoForm(getStoredOperator()));
  const [saoJoaoEntries, setSaoJoaoEntries] = useState(() => getStoredSaoJoaoQueue());
  const [saoJoaoSyncing, setSaoJoaoSyncing] = useState(false);
  const [saoJoaoImageLoadingKey, setSaoJoaoImageLoadingKey] = useState('');

  const canConfirmLocation = useMemo(
    () => Number.isFinite(draftPosition?.lat) && Number.isFinite(draftPosition?.lng),
    [draftPosition]
  );
  const needsLocationConfirmation = canConfirmLocation && (
    !confirmedPosition
    || confirmedPosition.lat !== draftPosition?.lat
    || confirmedPosition.lng !== draftPosition?.lng
  );

  const locationActionLabel = needsLocationConfirmation
    ? 'Confirmar local'
    : 'Capturar minha localização';
  const gpsAccuracyLabel = useMemo(() => formatAccuracy(gpsAccuracy), [gpsAccuracy]);
  const gpsAccuracyTone = useMemo(() => {
    if (!Number.isFinite(gpsAccuracy)) return '';
    if (gpsAccuracy <= 8) return 'good';
    if (gpsAccuracy <= 20) return 'fair';
    return 'poor';
  }, [gpsAccuracy]);

  const lastEntryTemplate = useMemo(() => {
    const operatorEntries = entries.filter((entry) => entry.OPERADOR_ID === activeOperator?.id);
    if (!operatorEntries.length) return null;

    return FAST_COPY_FIELDS.reduce((accumulator, field) => {
      accumulator[field] = operatorEntries[0][field] ?? '';
      return accumulator;
    }, {});
  }, [activeOperator, entries]);

  const operatorEntries = useMemo(
    () => entries.filter((entry) => entry.OPERADOR_ID === activeOperator?.id),
    [activeOperator, entries]
  );

  const pendingEntries = useMemo(
    () => entries.filter((entry) => entry.__syncStatus !== 'synced'),
    [entries]
  );
  const syncablePendingEntries = useMemo(
    () => pendingEntries.filter((entry) => isEntryReadyToSync(entry)),
    [pendingEntries]
  );
  const awaitingConfirmationEntries = useMemo(
    () => pendingEntries.filter((entry) => !isEntryReadyToSync(entry)),
    [pendingEntries]
  );
  const filteredOperatorEntries = useMemo(() => {
    if (queueFilter === 'ready') {
      return operatorEntries.filter((entry) => isEntryReadyToSync(entry));
    }
    if (queueFilter === 'waiting') {
      return operatorEntries.filter((entry) => !isEntryReadyToSync(entry));
    }
    return operatorEntries;
  }, [operatorEntries, queueFilter]);
  const requiredFieldsFilled = useMemo(
    () => {
      const baseFields = buildFormFocusSequence(form).filter((field) => !field.startsWith('LUMINARIA_'));
      const baseFilled = baseFields.filter((field) => String(form[field] ?? '').trim()).length;
      const luminariasFilled = (form.LUMINARIAS || []).reduce((total, item) => {
        const hasPotencia = String(item?.POTENCIA ?? '').trim().length > 0;
        const hasImagem = Boolean(item?.IMAGEM);
        return total + (hasPotencia ? 1 : 0) + (hasImagem ? 1 : 0);
      }, 0);
      return baseFilled + luminariasFilled;
    },
    [form]
  );
  const totalRequiredFields = buildFormFocusSequence(form).length;
  const remainingRequiredFields = Math.max(0, totalRequiredFields - requiredFieldsFilled);
  const formProgressPercent = Math.round((requiredFieldsFilled / Math.max(1, totalRequiredFields)) * 100);
  const requestAutofocus = useCallback(() => {
    setAutofocusRequestId((current) => current + 1);
  }, []);

  const guideElementByKey = useCallback((focusKey) => {
    const nextElement = formFieldRefs.current[focusKey];
    if (nextElement) {
      setGuidedFieldKey(focusKey);
      return true;
    }

    if (saveActionRef.current) {
      setGuidedFieldKey('SAVE_ACTION');
      return true;
    }

    return false;
  }, []);

  const focusFirstPendingFormField = useCallback(() => {
    const currentForm = formStateRef.current;
    const nextPendingField = buildFormFocusSequence(currentForm).find((focusKey) => !isFocusKeyFilled(currentForm, focusKey));
    guideElementByKey(nextPendingField || 'SAVE_ACTION');
  }, [guideElementByKey]);

  const focusNextFormField = useCallback((currentKey) => {
    const currentForm = formStateRef.current;
    const focusSequence = buildFormFocusSequence(currentForm);
    const currentIndex = focusSequence.indexOf(currentKey);
    const nextPendingField = focusSequence.find(
      (focusKey, index) => index > currentIndex && !isFocusKeyFilled(currentForm, focusKey)
    );
    guideElementByKey(nextPendingField || 'SAVE_ACTION');
  }, [guideElementByKey]);

  useEffect(() => {
    formStateRef.current = form;
  }, [form]);

  useEffect(() => {
    if (typeof window === 'undefined') return;
    window.localStorage.setItem(SAO_JOAO_QUEUE_STORAGE_KEY, JSON.stringify(saoJoaoEntries));
  }, [saoJoaoEntries]);

  useEffect(() => {
    if (!activeOperator) return undefined;
    let active = true;

    fetchFieldModules({ includeInactive: Boolean(activeOperator.can_export) }).then((modules) => {
      if (!active) return;
      setFieldModules(modules);
      if (activeModule && !modules.some((module) => module.slug === activeModule)) {
        setActiveModule(null);
        if (typeof window !== 'undefined') {
          window.localStorage.removeItem(MODULE_STORAGE_KEY);
        }
      }
    });

    return () => {
      active = false;
    };
  }, [activeModule, activeOperator]);

  useEffect(() => {
    if (!guidedFieldKey) return undefined;

    const timer = window.setTimeout(() => {
      setGuidedFieldKey('');
    }, 1800);

    return () => window.clearTimeout(timer);
  }, [guidedFieldKey]);

  useEffect(() => {
    if (!activeOperator) return;

    setForm((current) => ({
      ...current,
      OPERADOR: activeOperator.name,
    }));
    setSaoJoaoForm((current) => ({
      ...current,
      OPERADOR: activeOperator.name,
      OPERADOR_ID: activeOperator.id,
    }));
  }, [activeOperator]);

  useEffect(() => {
    let active = true;

    fetchAllowedOperators()
      .then((operators) => {
        if (!active) return;
        setAllowedOperators(buildMergedOperatorList(operators));
        setAuthLoading(false);
        setAuthError('');

        if (activeOperator && !operators.some((operator) => operator.id === activeOperator.id)) {
          setActiveOperator(null);
          if (typeof window !== 'undefined') {
            window.localStorage.removeItem(SESSION_STORAGE_KEY);
          }
        }
      })
      .catch((error) => {
        if (!active) return;
        const offlineOperators = buildMergedOperatorList();
        setAllowedOperators(offlineOperators);
        setAuthLoading(false);
        setAuthError(
          offlineOperators.length
            ? 'Sem conexão. Use um acesso offline já preparado neste aparelho.'
            : (error.message || 'Não foi possível carregar os usuários permitidos.')
        );
      });

    return () => {
      active = false;
    };
  }, [activeOperator]);

  useEffect(() => {
    if (!online) return;

    let active = true;

    fetchRpaOptions()
      .then((options) => {
        if (!active) return;
        setRpaOptions(options);
      })
      .catch(() => {
        if (!active) return;
        setRpaOptions([]);
      });

    return () => {
      active = false;
    };
  }, [online]);

  useEffect(() => {
    if (!activeOperator) return;
    if (activeModule !== 'implantacao-obra') return;
    if (!mapContainerRef.current || mapRef.current) return;

    const map = new maplibregl.Map({
      container: mapContainerRef.current,
      center: [-34.9000, -8.0500],
      zoom: 11.7,
      maxZoom: 20,
      dragRotate: false,
      pitchWithRotate: false,
      style: {
        version: 8,
        glyphs: 'https://demotiles.maplibre.org/font/{fontstack}/{range}.pbf',
        sources: {
          streets: {
            type: 'raster',
            tiles: MAP_TILE_URLS,
            tileSize: 256,
            attribution: '&copy; OpenStreetMap contributors &copy; CARTO',
          },
          recife: {
            type: 'geojson',
            data: RECIFE_BOUNDARY,
          },
          current: {
            type: 'geojson',
            data: createEmptyFeatureCollection(),
          },
        },
        layers: [
          { id: 'background', type: 'background', paint: { 'background-color': '#edf3fb' } },
          { id: 'streets', type: 'raster', source: 'streets' },
          {
            id: 'recife-fill',
            type: 'fill',
            source: 'recife',
            paint: { 'fill-color': '#1d4ed8', 'fill-opacity': 0.04 },
          },
          {
            id: 'recife-outline',
            type: 'line',
            source: 'recife',
            paint: { 'line-color': '#b91c1c', 'line-width': 3, 'line-dasharray': [3, 2] },
          },
          {
            id: 'current-point',
            type: 'circle',
            source: 'current',
            paint: {
              'circle-radius': 10,
              'circle-color': '#f59e0b',
              'circle-stroke-color': '#ffffff',
              'circle-stroke-width': 3,
            },
          },
        ],
      },
    });

    map.addControl(new maplibregl.NavigationControl({ showCompass: true }), 'top-right');
    mapRef.current = map;

    return () => {
      map.remove();
      mapRef.current = null;
      markerRef.current = null;
    };
  }, [activeModule, activeOperator]);

  useEffect(() => {
    const map = mapRef.current;
    if (!map) return;
    const source = map.getSource('current');
    if (!source) return;
    source.setData(buildPointCollection(draftPosition?.lat, draftPosition?.lng));
  }, [draftPosition]);

  useEffect(() => {
    if (!toast) return undefined;
    const timer = window.setTimeout(() => setToast(''), 2600);
    return () => window.clearTimeout(timer);
  }, [toast]);

  useEffect(() => {
    if (typeof window === 'undefined') return;
    window.localStorage.setItem(QUEUE_STORAGE_KEY, JSON.stringify(entries));
  }, [entries]);

  useEffect(() => {
    function handleOnline() {
      setOnline(true);
    }

    function handleOffline() {
      setOnline(false);
    }

    window.addEventListener('online', handleOnline);
    window.addEventListener('offline', handleOffline);

    return () => {
      window.removeEventListener('online', handleOnline);
      window.removeEventListener('offline', handleOffline);
    };
  }, []);

  useEffect(() => {
    if (step !== 'form' || !formSectionRef.current) return;
    setGuidedFieldKey('');
  }, [step, confirmedPosition]);

  useEffect(() => {
    if (step !== 'form' || !autofocusRequestId || handledAutofocusRequestRef.current === autofocusRequestId) return;
    handledAutofocusRequestRef.current = autofocusRequestId;

    const timer = window.setTimeout(() => {
      focusFirstPendingFormField();
    }, 120);

    return () => window.clearTimeout(timer);
  }, [autofocusRequestId, focusFirstPendingFormField, form, step]);

  const focusMapOnPosition = useCallback((lat, lng) => {
    const map = mapRef.current;
    if (!map || !Number.isFinite(lat) || !Number.isFinite(lng)) return;

    map.easeTo({ center: [lng, lat], zoom: 18.2, duration: 900 });
    if (!markerRef.current) {
      markerRef.current = new maplibregl.Marker({ draggable: true, color: '#f59e0b' })
        .setLngLat([lng, lat])
        .addTo(map);
      markerRef.current.on('dragend', () => {
        const { lat: markerLat, lng: markerLng } = markerRef.current.getLngLat();
        setDraftPosition({ lat: markerLat, lng: markerLng });
        setConfirmedPosition(null);
        setManualCoordinateValue(`${formatCoordinate(markerLat)}, ${formatCoordinate(markerLng)}`);
        setManualCoordinateError('');
        setToast('Pin movido. Confirme de novo para atualizar o ponto.');
      });
      return;
    }

    markerRef.current.setLngLat([lng, lat]);
  }, []);

  const handleCaptureLocation = useCallback(() => {
    if (!navigator.geolocation) {
      setCaptureError('Seu dispositivo não suporta geolocalização.');
      return;
    }

    setGeoStatus('Captando localização...');
    setCaptureError('');
    setManualCoordinateError('');

    navigator.geolocation.getCurrentPosition(
      (position) => {
        const lat = Number(position.coords.latitude);
        const lng = Number(position.coords.longitude);
        const accuracy = Number(position.coords.accuracy);
        const next = { lat, lng };
        setDraftPosition(next);
        setConfirmedPosition(null);
        setGpsAccuracy(Number.isFinite(accuracy) ? accuracy : null);
        setGeoStatus('Localização captada. Se precisar, arraste o pin.');
        setToast('Ajuste o pin e confirme.');
        focusMapOnPosition(lat, lng);
      },
      (error) => {
        setGeoStatus('Não foi possível captar a localização.');
        setCaptureError(error.message || 'Verifique as permissões de localização do celular.');
      },
      { enableHighAccuracy: true, timeout: 15000, maximumAge: 0 }
    );
  }, [focusMapOnPosition]);

  const applyManualCoordinates = useCallback((rawValue) => {
    const parsed = parseManualCoordinates(rawValue);
    if (!parsed) {
      setManualCoordinateError('Cole latitude e longitude no formato -8.052240, -34.928610.');
      return false;
    }

    setCaptureError('');
    setManualCoordinateError('');
    setDraftPosition(parsed);
    setConfirmedPosition(null);
    setGpsAccuracy(null);
    setManualCoordinateValue(`${formatCoordinate(parsed.lat)}, ${formatCoordinate(parsed.lng)}`);
    setGeoStatus('Coordenadas aplicadas. Se precisar, arraste o pin.');
    setToast('Coordenadas aplicadas. Confirme o local.');
    focusMapOnPosition(parsed.lat, parsed.lng);
    return true;
  }, [focusMapOnPosition]);

  const handleManualCoordinateSubmit = useCallback((event) => {
    event.preventDefault();
    applyManualCoordinates(manualCoordinateValue);
  }, [applyManualCoordinates, manualCoordinateValue]);

  const handleManualCoordinatePaste = useCallback((event) => {
    const pastedText = event.clipboardData?.getData('text') ?? '';
    if (!pastedText.trim()) return;

    setManualCoordinateValue(pastedText);
    if (parseManualCoordinates(pastedText)) {
      event.preventDefault();
      applyManualCoordinates(pastedText);
    }
  }, [applyManualCoordinates]);

  const handleConfirmLocation = useCallback(async () => {
    if (!canConfirmLocation) return;
    setConfirmedPosition(draftPosition);
    setForm((current) => ({
      ...current,
      LATITUDE: formatCoordinate(draftPosition.lat),
      LONGITUDE: formatCoordinate(draftPosition.lng),
      ATUALIZACAO: new Date().toISOString().slice(0, 10),
    }));
    setStep('form');
    setToast('Local confirmado. Agora preencha o ponto.');
    requestAutofocus();

    try {
      const context = await fetchLocationContext(draftPosition.lat, draftPosition.lng);
      setForm((current) => ({
        ...current,
        LATITUDE: formatCoordinate(draftPosition.lat),
        LONGITUDE: formatCoordinate(draftPosition.lng),
        ATUALIZACAO: new Date().toISOString().slice(0, 10),
        ENDERECO: current.ENDERECO || context.endereco || '',
        BAIRRO: current.BAIRRO || context.bairro || '',
        RPA: current.RPA || context.rpa || '',
      }));
      requestAutofocus();
    } catch {
      // Autofill is best-effort; the form remains usable even without context.
    }
  }, [canConfirmLocation, draftPosition, requestAutofocus]);

  const handleLocationAction = useCallback(() => {
    if (needsLocationConfirmation) {
      handleConfirmLocation();
      return;
    }

    handleCaptureLocation();
  }, [needsLocationConfirmation, handleCaptureLocation, handleConfirmLocation]);

  const handleRecenterGps = useCallback(() => {
    handleCaptureLocation();
  }, [handleCaptureLocation]);

  const handleChangeField = useCallback((field, value) => {
    if (field === 'QTDE') {
      const nextCount = clampLuminaireCount(value);
      setForm((current) => ({
        ...current,
        QTDE: String(nextCount),
        LUMINARIAS: normalizeLuminaireItems(current.LUMINARIAS, nextCount),
      }));
      return;
    }
    setForm((current) => ({
      ...current,
      [field]: value,
      ...(field === 'MOTIVO_IMPLANTACAO' && value !== 'OBRA' ? { OBRA_NOME: '' } : {}),
    }));
  }, []);

  const handleLuminairePotenciaChange = useCallback((index, value) => {
    setForm((current) => ({
      ...current,
      LUMINARIAS: current.LUMINARIAS.map((item, itemIndex) => (
        itemIndex === index ? { ...item, POTENCIA: value } : item
      )),
    }));
  }, []);

  const handleImageSelection = useCallback(async (index, event) => {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file) return;

    try {
      const imagePayload = await prepareImagePayload(file);
      setForm((current) => ({
        ...current,
        LUMINARIAS: current.LUMINARIAS.map((item, itemIndex) => (
          itemIndex === index ? { ...item, IMAGEM: imagePayload } : item
        )),
      }));
      setToast('Foto adicionada ao ponto.');
      window.setTimeout(() => focusNextFormField(`LUMINARIA_${index + 1}_IMAGEM`), 0);
    } catch (error) {
      setToast(error.message || 'Não foi possível preparar a foto.');
    }
  }, [focusNextFormField]);

  const handleRemoveImage = useCallback((index) => {
    setForm((current) => ({
      ...current,
      LUMINARIAS: current.LUMINARIAS.map((item, itemIndex) => (
        itemIndex === index ? { ...item, IMAGEM: null } : item
      )),
    }));
    setToast('Foto removida deste ponto.');
  }, []);

  const handleAccessFieldChange = useCallback((field, value) => {
    setAccessForm((current) => ({ ...current, [field]: value }));
    setAuthError('');
    setAuthInfo('');
  }, []);

  const handleToggleOfflineRpa = useCallback((rpa) => {
    setAccessForm((current) => {
      const currentValues = current.offlineRpas || [];
      const exists = currentValues.includes(rpa);
      if (exists) {
        return { ...current, offlineRpas: currentValues.filter((item) => item !== rpa) };
      }
      if (currentValues.length >= MAX_OFFLINE_RPAS) {
        setAuthError(`Escolha no máximo ${MAX_OFFLINE_RPAS} RPAs para o preparo offline.`);
        setAuthInfo('');
        return current;
      }
      return { ...current, offlineRpas: [...currentValues, rpa] };
    });
  }, []);

  const handleManagerUserFieldChange = useCallback((field, value) => {
    setManagerUserForm((current) => ({ ...current, [field]: value }));
  }, []);

  const handleManagerFilterChange = useCallback((field, value) => {
    setManagerTableFilters((current) => ({ ...current, [field]: value }));
  }, []);

  const handleLogin = useCallback(async (event) => {
    event.preventDefault();

    try {
      const trimmedAccessCode = accessForm.accessCode.trim();
      const selectedOfflineRpas = accessForm.offlineRpas || [];
      const loggedOperator = online
        ? await loginFieldOperator(accessForm.operatorId, trimmedAccessCode)
        : await loginOfflineOperator(accessForm.operatorId, trimmedAccessCode);

      if (online && selectedOfflineRpas.length) {
        setOfflinePrepLoading(true);
        const cachedTiles = await prepareOfflineAccess(loggedOperator, trimmedAccessCode, selectedOfflineRpas);
        setAuthInfo(`Acesso offline pronto para ${selectedOfflineRpas.join(', ')}. ${cachedTiles} tile(s) preparados neste aparelho.`);
      } else {
        setAuthInfo('');
      }

      setActiveOperator(loggedOperator);
      setActiveModule(null);
      setAuthError('');
      setAccessForm({ operatorId: String(loggedOperator.id), accessCode: '', offlineRpas: selectedOfflineRpas });
      if (typeof window !== 'undefined') {
        window.localStorage.setItem(
          SESSION_STORAGE_KEY,
          JSON.stringify({
            id: loggedOperator.id,
            name: loggedOperator.name,
            can_export: loggedOperator.can_export,
          })
        );
        window.localStorage.removeItem(MODULE_STORAGE_KEY);
      }
      setToast(`Acesso liberado para ${loggedOperator.name}.`);
    } catch (error) {
      setAuthError(error.message || 'Acesso nao autorizado. Confira usuario e codigo.');
      setAuthInfo('');
    } finally {
      setOfflinePrepLoading(false);
    }
  }, [accessForm, online]);

  const handleLogout = useCallback(() => {
    setActiveOperator(null);
    setActiveModule(null);
    setAccessForm({ operatorId: '', accessCode: '', offlineRpas: [] });
    setAuthError('');
    setAuthInfo('');
    setStep('location');
    setDraftPosition(null);
    setConfirmedPosition(null);
    setEditingEntryId(null);
    setNextPointReady(false);
    setGpsAccuracy(null);
    setManualCoordinateValue('');
    setManualCoordinateError('');
    setForm(INITIAL_FORM);
    setToast('');
    if (markerRef.current) {
      markerRef.current.remove();
      markerRef.current = null;
    }
    if (typeof window !== 'undefined') {
      window.localStorage.removeItem(SESSION_STORAGE_KEY);
      window.localStorage.removeItem(MODULE_STORAGE_KEY);
    }
  }, []);

  const handleSelectModule = useCallback((moduleSlug) => {
    setActiveModule(moduleSlug);
    if (typeof window !== 'undefined') {
      window.localStorage.setItem(MODULE_STORAGE_KEY, moduleSlug);
    }
    setToast(moduleSlug === 'sao-joao' ? 'Módulo São João aberto.' : 'Módulo Implantação/Obra aberto.');
  }, []);

  const handleBackToModules = useCallback(() => {
    setActiveModule(null);
    if (typeof window !== 'undefined') {
      window.localStorage.removeItem(MODULE_STORAGE_KEY);
    }
  }, []);

  const handleSaoJoaoFieldChange = useCallback((field, value) => {
    setSaoJoaoForm((current) => ({
      ...current,
      [field]: value,
    }));
  }, []);

  const handleSaoJoaoAnswerChange = useCallback((field, value) => {
    setSaoJoaoForm((current) => ({
      ...current,
      RESPOSTAS: {
        ...current.RESPOSTAS,
        [field]: value,
      },
    }));
  }, []);

  const handleSaoJoaoImageSelection = useCallback(async (fieldKey, event) => {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file) return;

    try {
      setSaoJoaoImageLoadingKey(fieldKey);
      const imagePayload = await prepareImagePayload(file);
      setSaoJoaoForm((current) => ({
        ...current,
        IMAGENS: {
          ...current.IMAGENS,
          [fieldKey]: imagePayload,
        },
      }));
      setToast('Foto adicionada ao checklist.');
    } catch (error) {
      setToast(error.message || 'Não foi possível preparar a foto.');
    } finally {
      setSaoJoaoImageLoadingKey('');
    }
  }, []);

  const handleRemoveSaoJoaoImage = useCallback((fieldKey) => {
    setSaoJoaoForm((current) => {
      const nextImages = { ...current.IMAGENS };
      delete nextImages[fieldKey];
      return { ...current, IMAGENS: nextImages };
    });
  }, []);

  const handleCaptureSaoJoaoLocation = useCallback(() => {
    if (!navigator.geolocation) {
      setToast('GPS indisponível neste aparelho.');
      return;
    }

    navigator.geolocation.getCurrentPosition(
      (position) => {
        const lat = formatCoordinate(position.coords.latitude);
        const lng = formatCoordinate(position.coords.longitude);
        setSaoJoaoForm((current) => ({
          ...current,
          LATITUDE: lat,
          LONGITUDE: lng,
        }));
        setToast('Localização adicionada ao checklist.');
      },
      () => setToast('Não foi possível capturar o GPS agora. Você pode colar as coordenadas.'),
      { enableHighAccuracy: true, timeout: 12000, maximumAge: 0 }
    );
  }, []);

  const handlePasteSaoJoaoCoordinates = useCallback((value) => {
    const parsed = parseManualCoordinates(value);
    if (!parsed) {
      setToast('Cole latitude e longitude em um formato válido.');
      return;
    }
    setSaoJoaoForm((current) => ({
      ...current,
      LATITUDE: formatCoordinate(parsed.lat),
      LONGITUDE: formatCoordinate(parsed.lng),
    }));
  }, []);

  const handleSaveSaoJoaoEntry = useCallback(() => {
    const nextEntry = {
      ...saoJoaoForm,
      CLIENT_UUID: crypto.randomUUID(),
      OPERADOR: activeOperator?.name || saoJoaoForm.OPERADOR,
      OPERADOR_ID: activeOperator?.id || saoJoaoForm.OPERADOR_ID,
      __id: crypto.randomUUID(),
      __syncStatus: 'pending',
      __createdAt: new Date().toISOString(),
    };

    setSaoJoaoEntries((current) => [nextEntry, ...current]);
    setSaoJoaoForm(createInitialSaoJoaoForm(activeOperator));
    setToast('Checklist salvo na fila local.');
  }, [activeOperator, saoJoaoForm]);

  const handleSyncSaoJoaoEntries = useCallback(async () => {
    const pending = saoJoaoEntries.filter((entry) => entry.__syncStatus !== 'synced');
    if (!pending.length) {
      setToast('Não há checklist pendente para sincronizar.');
      return;
    }
    if (!online) {
      setToast('Sem conexão. O checklist continua salvo no aparelho.');
      return;
    }

    try {
      setSaoJoaoSyncing(true);
      const result = await syncSaoJoaoEntries(pending);
      const syncedIds = new Set(result.client_uuids || []);
      setSaoJoaoEntries((current) => current.filter((entry) => !syncedIds.has(entry.CLIENT_UUID)));
      setToast(`${syncedIds.size} checklist(s) do São João sincronizado(s).`);
    } catch (error) {
      setToast(error.message || 'Falha ao sincronizar o checklist do São João.');
    } finally {
      setSaoJoaoSyncing(false);
    }
  }, [online, saoJoaoEntries]);

  const handleRemoveSaoJoaoEntry = useCallback((entryId) => {
    setSaoJoaoEntries((current) => current.filter((entry) => entry.__id !== entryId));
  }, []);

  const handleStartNextPoint = useCallback(() => {
    const seedForm = lastEntryTemplate
      ? buildNextPointForm(lastEntryTemplate, activeOperator?.name || '')
      : buildNextPointForm({}, activeOperator?.name || '');

    setEditingEntryId(null);
    setConfirmedPosition(null);
    setDraftPosition(null);
    setGpsAccuracy(null);
    setManualCoordinateValue('');
    setManualCoordinateError('');
    setForm(seedForm);
    setStep('location');
    setNextPointReady(false);
    if (markerRef.current) {
      markerRef.current.remove();
      markerRef.current = null;
    }
    handleCaptureLocation();
  }, [activeOperator, handleCaptureLocation, lastEntryTemplate]);

  const performSyncEntries = useCallback(async (targetEntries) => {
    if (!targetEntries.length || syncing || !online) return;
    setSyncing(true);
    try {
      const payloadEntries = targetEntries.map((entry) => ({
        ...normalizeEntryShape(entry),
        CLIENT_UUID: entry.CLIENT_UUID,
        OPERADOR_ID: entry.OPERADOR_ID,
        OPERADOR: entry.OPERADOR,
        ID_CADASTRO: entry.ID_CADASTRO,
        TIPO_IMPLANTACAO: entry.TIPO_IMPLANTACAO,
        MOTIVO_IMPLANTACAO: entry.MOTIVO_IMPLANTACAO,
        OBRA_NOME: entry.MOTIVO_IMPLANTACAO === 'OBRA' ? entry.OBRA_NOME : '',
        IMPLANTACAO_CONCLUIDA: entry.IMPLANTACAO_CONCLUIDA,
        RPA: entry.RPA,
        LOCALIZACA: entry.LOCALIZACA,
        ENDERECO: entry.ENDERECO,
        BARRAMENTO: entry.BARRAMENTO,
        LATITUDE: entry.LATITUDE,
        LONGITUDE: entry.LONGITUDE,
        BAIRRO: entry.BAIRRO,
        MEDICAO: entry.MEDICAO,
        TIPO_LUMIN: entry.TIPO_LUMIN,
        TIPO_DE_PO: entry.TIPO_DE_PO,
        TIPO_LAMPA: entry.TIPO_LAMPA,
        QTDE: entry.QTDE,
        POTENCIA: computeLuminairePotenciaTotal(entry.LUMINARIAS ?? []),
        LUMINARIAS: normalizeLuminaireItems(entry.LUMINARIAS, entry.QTDE),
        PERDAS: entry.PERDAS,
        TOTAL_CARGA: entry.TOTAL_CARGA,
        CONSUMO_kW: entry.CONSUMO_kW,
        CONSUMO_kW_MES: entry.CONSUMO_kW_MES,
        ATUALIZACAO: entry.ATUALIZACAO,
        BLOQ_AMP: entry.BLOQ_AMP,
      }));

      const result = await syncFieldEntries(payloadEntries);
      const syncedIds = new Set(result.client_uuids || []);
      setEntries((current) => current.map((entry) => (
        syncedIds.has(entry.CLIENT_UUID)
          ? {
              ...entry,
              __syncStatus: 'synced',
              __syncedAt: new Date().toISOString(),
              __removing: true,
            }
          : entry
      )));
      setToast(`${syncedIds.size} registro(s) sincronizado(s) com o banco.`);
      window.setTimeout(() => {
        setEntries((current) => current.filter((entry) => !syncedIds.has(entry.CLIENT_UUID)));
      }, 1200);
    } catch (error) {
      setEntries((current) => current.map((entry) => (
        targetEntries.some((pending) => pending.CLIENT_UUID === entry.CLIENT_UUID)
          ? { ...entry, __syncStatus: 'error' }
          : entry
      )));
      setToast(error.message || 'Falha na sincronização. Os dados continuam salvos no aparelho.');
    } finally {
      setSyncing(false);
    }
  }, [online, syncing]);

  const handleSyncEntries = useCallback(async () => {
    if (!syncablePendingEntries.length) {
      setToast('Confirme o local e marque implantação concluída = SIM antes de sincronizar.');
      return;
    }

    await performSyncEntries(syncablePendingEntries);
  }, [performSyncEntries, syncablePendingEntries]);

  const loadManagerSummary = useCallback(async () => {
    if (!activeOperator?.can_export) return;
    setManagerLoading(true);
    try {
      const summary = await fetchSubmissionSummary();
      setManagerSummary(summary);
    } catch (error) {
      setToast(error.message || 'Não foi possível carregar o resumo gerencial.');
    } finally {
      setManagerLoading(false);
    }
  }, [activeOperator, setToast]);

  const loadManagerRows = useCallback(async () => {
    if (!activeOperator?.can_export) return;
    setManagerRowsLoading(true);
    try {
      const rows = await fetchSubmissionExport('all');
      setManagerRows(rows);
    } catch (error) {
      setToast(error.message || 'Não foi possível carregar os pontos registrados.');
    } finally {
      setManagerRowsLoading(false);
    }
  }, [activeOperator, setToast]);

  useEffect(() => {
    if (!activeOperator?.can_export) return;
    loadManagerSummary();
    loadManagerRows();
  }, [activeOperator, loadManagerSummary, loadManagerRows]);

  const handleStartManagerEdit = useCallback((row) => {
    const clientUuid = row.client_uuid || row.CLIENT_UUID || '';
    if (!clientUuid) {
      setToast('Não foi possível identificar este ponto para edição. Atualize a lista e tente novamente.');
      return;
    }
    setManagerEditingClientUuid(clientUuid);
    setManagerEditDraft({
      RPA: row.rpa || '',
      ENDERECO: row.endereco || '',
      BAIRRO: row.bairro || '',
      TIPO_IMPLANTACAO: row.tipo_implantacao || '',
      MOTIVO_IMPLANTACAO: row.motivo_implantacao || '',
      OBRA_NOME: row.obra_nome || '',
      IMPLANTACAO_CONCLUIDA: row.implantacao_concluida || '',
      TIPO_DE_PO: row.tipo_de_po || '',
    });
  }, []);

  const handleManagerEditFieldChange = useCallback((field, value) => {
    setManagerEditDraft((current) => ({
      ...current,
      [field]: value,
      ...(field === 'MOTIVO_IMPLANTACAO' && value !== 'OBRA' ? { OBRA_NOME: '' } : {}),
    }));
  }, []);

  const handleSaveManagerEdit = useCallback(async (clientUuid) => {
    if (!activeOperator?.can_export) return;
    if (!clientUuid) {
      setToast('Não foi possível identificar este ponto para edição. Atualize a lista e tente novamente.');
      return;
    }
    if (!managerUserForm.managerAccessCode.trim()) {
      setToast('Informe o código gerencial antes de editar pontos.');
      return;
    }
    try {
      setManagerRowActionId(clientUuid);
      await updateFieldSubmission(clientUuid, {
        manager_operator_id: activeOperator.id,
        manager_access_code: managerUserForm.managerAccessCode.trim(),
        RPA: managerEditDraft.RPA,
        ENDERECO: managerEditDraft.ENDERECO,
        BAIRRO: managerEditDraft.BAIRRO,
        TIPO_IMPLANTACAO: managerEditDraft.TIPO_IMPLANTACAO,
        MOTIVO_IMPLANTACAO: managerEditDraft.MOTIVO_IMPLANTACAO,
        OBRA_NOME: managerEditDraft.MOTIVO_IMPLANTACAO === 'OBRA' ? managerEditDraft.OBRA_NOME : '',
        IMPLANTACAO_CONCLUIDA: managerEditDraft.IMPLANTACAO_CONCLUIDA,
        TIPO_DE_PO: managerEditDraft.TIPO_DE_PO,
      });
      setManagerEditingClientUuid('');
      setManagerEditDraft({});
      await loadManagerRows();
      await loadManagerSummary();
      setToast('Ponto atualizado.');
    } catch (error) {
      setToast(error.message || 'Não foi possível editar o ponto.');
    } finally {
      setManagerRowActionId('');
    }
  }, [activeOperator, loadManagerRows, loadManagerSummary, managerEditDraft, managerUserForm.managerAccessCode]);

  const handleDeleteManagerRow = useCallback(async (row) => {
    if (!activeOperator?.can_export) return;
    if (!managerUserForm.managerAccessCode.trim()) {
      setToast('Informe o código gerencial antes de excluir pontos.');
      return;
    }
    const clientUuid = row.client_uuid || row.CLIENT_UUID || '';
    if (!clientUuid) {
      setToast('Não foi possível identificar este ponto para exclusão. Atualize a lista e tente novamente.');
      return;
    }
    const confirmed = window.confirm(`Excluir o ponto registrado por ${row.operador || 'campo'}?`);
    if (!confirmed) return;

    try {
      setManagerRowActionId(clientUuid);
      await deleteFieldSubmission(clientUuid, {
        manager_operator_id: activeOperator.id,
        manager_access_code: managerUserForm.managerAccessCode.trim(),
      });
      await loadManagerRows();
      await loadManagerSummary();
      setToast('Ponto excluído.');
    } catch (error) {
      setToast(error.message || 'Não foi possível excluir o ponto.');
    } finally {
      setManagerRowActionId('');
    }
  }, [activeOperator, loadManagerRows, loadManagerSummary, managerUserForm.managerAccessCode]);

  const handleToggleManagerModule = useCallback(async (module) => {
    if (!activeOperator?.can_export) return;
    if (!managerUserForm.managerAccessCode.trim()) {
      setToast('Informe o código gerencial antes de alterar módulos.');
      return;
    }

    try {
      setManagerModuleActionSlug(module.slug);
      await updateFieldModule(module.slug, {
        manager_operator_id: activeOperator.id,
        manager_access_code: managerUserForm.managerAccessCode.trim(),
        active: !module.active,
      });
      const modules = await fetchFieldModules({ includeInactive: true });
      setFieldModules(modules);
      setToast(`${module.name} foi ${module.active ? 'desativado' : 'ativado'}.`);
    } catch (error) {
      setToast(error.message || 'Não foi possível alterar o módulo.');
    } finally {
      setManagerModuleActionSlug('');
    }
  }, [activeOperator, managerUserForm.managerAccessCode]);

  const renderManagerTableCell = useCallback((field, exportRow, isEditing) => {
    const value = exportRow[field];

    if (isEditing && MANAGER_EDITABLE_FIELDS.has(field)) {
      if (field === 'MOTIVO_IMPLANTACAO') {
        return (
          <select value={managerEditDraft.MOTIVO_IMPLANTACAO || ''} onChange={(event) => handleManagerEditFieldChange('MOTIVO_IMPLANTACAO', event.target.value)}>
            <option value="">Motivo</option>
            <option value="OBRA">OBRA</option>
            <option value="EXPANSÃO">EXPANSÃO</option>
          </select>
        );
      }

      if (field === 'OBRA_NOME') {
        return (
          <select
            value={managerEditDraft.OBRA_NOME || ''}
            onChange={(event) => handleManagerEditFieldChange('OBRA_NOME', event.target.value)}
            disabled={managerEditDraft.MOTIVO_IMPLANTACAO !== 'OBRA'}
          >
            <option value="">Obra</option>
            <option value="PEDONAL CAXANGÁ">PEDONAL CAXANGÁ</option>
            <option value="PEDONAL ZONA OESTE">PEDONAL ZONA OESTE</option>
          </select>
        );
      }

      if (field === 'TIPO_DE_PO') {
        return (
          <select value={managerEditDraft.TIPO_DE_PO || ''} onChange={(event) => handleManagerEditFieldChange('TIPO_DE_PO', event.target.value)}>
            <option value="">Tipo</option>
            <option value="CONCRETO">CONCRETO</option>
            <option value="METÁLICO">METÁLICO</option>
            <option value="FIBRA">FIBRA</option>
            <option value="GIRAFA">GIRAFA</option>
          </select>
        );
      }

      if (field === 'IMPLANTACAO_CONCLUIDA') {
        return (
          <select value={managerEditDraft.IMPLANTACAO_CONCLUIDA || ''} onChange={(event) => handleManagerEditFieldChange('IMPLANTACAO_CONCLUIDA', event.target.value)}>
            <option value="">Status</option>
            <option value="SIM">SIM</option>
            <option value="NÃO">NÃO</option>
          </select>
        );
      }

      return (
        <input
          value={managerEditDraft[field] || ''}
          placeholder={MANAGER_FIELD_LABELS[field] || field}
          onChange={(event) => handleManagerEditFieldChange(field, event.target.value)}
        />
      );
    }

    const formattedValue = formatManagerCellValue(value);
    if (formattedValue === '-') {
      return <span className="manager-table-empty">-</span>;
    }

    if (isManagerImageLinkField(field)) {
      const links = String(value)
        .split('\n')
        .map((item) => item.trim())
        .filter(Boolean);

      if (!links.length) {
        return <span className="manager-table-empty">-</span>;
      }

      return (
        <div className="manager-table-links">
          {links.map((link, index) => (
            <a key={`${field}-${link}`} href={link} target="_blank" rel="noreferrer">
              {links.length > 1 ? `Imagem ${index + 1}` : 'Abrir imagem'}
            </a>
          ))}
        </div>
      );
    }

    return formattedValue;
  }, [handleManagerEditFieldChange, managerEditDraft]);

  const renderManagerDetailGroup = useCallback((fields, exportRow, isEditing) => (
    <div className="manager-detail-grid">
      {fields.map((field) => (
        <div key={field} className="manager-detail-item">
          <span className="manager-detail-label">{MANAGER_FIELD_LABELS[field] || field}</span>
          <div className="manager-detail-value">
            {renderManagerTableCell(field, exportRow, isEditing)}
          </div>
        </div>
      ))}
    </div>
  ), [renderManagerTableCell]);

  const managerTableRows = useMemo(() => managerRows
    .map((row, index) => ({
      row,
      index,
      exportRow: buildExportRow(row),
      clientUuid: row.client_uuid || row.CLIENT_UUID || '',
    }))
    .filter(({ exportRow }) => MANAGER_COMPACT_TABLE_FIELDS.every((field) => {
      const filter = (managerTableFilters[field] || '').trim().toLowerCase();
      if (!filter) return true;
      return formatManagerCellValue(exportRow[field]).toLowerCase().includes(filter);
    })), [managerRows, managerTableFilters]);

  const handleCreateManagerUser = useCallback(async (event) => {
    event.preventDefault();
    if (!activeOperator?.can_export) return;

    if (!managerUserForm.managerAccessCode.trim() || !managerUserForm.name.trim() || !managerUserForm.accessCode.trim()) {
      setToast('Informe o código gerencial, o nome e o código de acesso do novo usuário.');
      return;
    }

    try {
      setManagerUserLoading(true);
      const createdUser = await createFieldOperator(activeOperator.id, managerUserForm.managerAccessCode.trim(), {
        name: managerUserForm.name.trim(),
        accessCode: managerUserForm.accessCode.trim(),
        canExport: managerUserForm.canExport,
      });
      const operators = await fetchAllowedOperators();
      setAllowedOperators(operators);
      setManagerUserForm((current) => ({ ...current, name: '', accessCode: '', canExport: false }));
      setToast(`${createdUser.name} foi adicionado com sucesso.`);
    } catch (error) {
      setToast(error.message || 'Não foi possível cadastrar o usuário.');
    } finally {
      setManagerUserLoading(false);
    }
  }, [activeOperator, managerUserForm]);

  const handleDeactivateManagerUser = useCallback(async (operator) => {
    if (!activeOperator?.can_export) return;
    if (operator.id === activeOperator.id) {
      setToast('Você não pode desativar o próprio acesso gerencial por aqui.');
      return;
    }

    const confirmed = window.confirm(`Desativar o acesso de ${operator.name}?`);
    if (!confirmed) return;

    try {
      setManagerUserActionId(operator.id);
      if (!managerUserForm.managerAccessCode.trim()) {
        setToast('Informe o código gerencial antes de desativar usuários.');
        return;
      }

      await deactivateFieldOperator(activeOperator.id, managerUserForm.managerAccessCode.trim(), operator.id);
      const operators = await fetchAllowedOperators();
      setAllowedOperators(operators);
      setToast(`${operator.name} foi desativado com sucesso.`);
    } catch (error) {
      setToast(error.message || 'Não foi possível desativar o usuário.');
    } finally {
      setManagerUserActionId(null);
    }
  }, [activeOperator]);

  const handleApplyQuickOption = useCallback((field, value) => {
    if (field === 'QTDE') {
      const nextCount = clampLuminaireCount(value);
      setForm((current) => ({
        ...current,
        QTDE: String(nextCount),
        LUMINARIAS: normalizeLuminaireItems(current.LUMINARIAS, nextCount),
      }));
      window.setTimeout(() => focusNextFormField(field), 0);
      return;
    }

    setForm((current) => ({
      ...current,
      [field]: value,
      ...(field === 'MOTIVO_IMPLANTACAO' && value !== 'OBRA' ? { OBRA_NOME: '' } : {}),
    }));
    window.setTimeout(() => focusNextFormField(field), 0);
  }, [focusNextFormField]);

  const handleFieldKeyDown = useCallback((field, event) => {
    if (event.key !== 'Enter') return;
    event.preventDefault();
    focusNextFormField(field);
  }, [focusNextFormField]);

  const handleLuminaireFieldKeyDown = useCallback((focusKey, event) => {
    if (event.key !== 'Enter') return;
    event.preventDefault();
    focusNextFormField(focusKey);
  }, [focusNextFormField]);

  const handleReuseLastEntry = useCallback(() => {
    if (!lastEntryTemplate) {
      setToast('Salve um primeiro registro para reaproveitar o padrão.');
      return;
    }

    setForm((current) => ({
      ...current,
      ...lastEntryTemplate,
      LATITUDE: current.LATITUDE,
      LONGITUDE: current.LONGITUDE,
      ENDERECO: current.ENDERECO,
      ID_CADASTRO: current.ID_CADASTRO,
      BARRAMENTO: current.BARRAMENTO,
      ATUALIZACAO: new Date().toISOString().slice(0, 10),
    }));
    setToast('Último padrão aplicado. Ajuste apenas o que mudar.');
    requestAutofocus();
  }, [lastEntryTemplate, requestAutofocus]);

  const handleClearOptionalFields = useCallback(() => {
    setForm((current) => ({
      ...current,
      ID_CADASTRO: '',
      LOCALIZACA: '',
      BARRAMENTO: '',
      LUMINARIAS: current.LUMINARIAS.map((item) => ({ ...item, IMAGEM: null })),
      MEDICAO: '',
      PERDAS: '',
      TOTAL_CARGA: '',
      CONSUMO_kW: '',
      CONSUMO_kW_MES: '',
      BLOQ_AMP: '',
    }));
    setToast('Campos opcionais limpos.');
    requestAutofocus();
  }, [requestAutofocus]);

  const handleEditEntry = useCallback((entry) => {
    const lat = toNumber(entry.LATITUDE);
    const lng = toNumber(entry.LONGITUDE);

    setEditingEntryId(entry.__id);
    setNextPointReady(false);
    setGpsAccuracy(null);
    setManualCoordinateError('');
    setForm(normalizeEntryShape({
      ...INITIAL_FORM,
      ...entry,
      OPERADOR: activeOperator?.name || entry.OPERADOR || '',
      ATUALIZACAO: entry.ATUALIZACAO || new Date().toISOString().slice(0, 10),
    }));
    if (Number.isFinite(lat) && Number.isFinite(lng)) {
      const nextPosition = { lat, lng };
      setManualCoordinateValue(`${formatCoordinate(lat)}, ${formatCoordinate(lng)}`);
      setDraftPosition(nextPosition);
      setConfirmedPosition(nextPosition);
      focusMapOnPosition(lat, lng);
    } else {
      setManualCoordinateValue('');
    }
    setStep('form');
    setToast('Ponto carregado. Ajuste o que faltar antes de enviar.');
    requestAutofocus();
  }, [activeOperator, focusMapOnPosition, requestAutofocus]);

  const handleSaveEntry = useCallback(() => {
    if (!confirmedPosition) {
      setToast('Confirme o local antes de salvar.');
      setStep('location');
      return;
    }

    const normalizedLuminarias = normalizeLuminaireItems(form.LUMINARIAS, form.QTDE);
    const nextEntry = {
      ...form,
      CLIENT_UUID: editingEntryId
        ? (entries.find((entry) => entry.__id === editingEntryId)?.CLIENT_UUID || crypto.randomUUID())
        : crypto.randomUUID(),
      OPERADOR: activeOperator?.name || form.OPERADOR,
      OPERADOR_ID: activeOperator?.id || null,
      QTDE: String(clampLuminaireCount(form.QTDE)),
      LUMINARIAS: normalizedLuminarias,
      LATITUDE: formatCoordinate(toNumber(form.LATITUDE)),
      LONGITUDE: formatCoordinate(toNumber(form.LONGITUDE)),
      POTENCIA: computeLuminairePotenciaTotal(normalizedLuminarias),
      __syncStatus: isEntryReadyToSync(form) ? 'pending' : 'draft',
      __createdAt: editingEntryId
        ? (entries.find((entry) => entry.__id === editingEntryId)?.__createdAt || new Date().toISOString())
        : new Date().toISOString(),
    };
    const nextPointSeed = buildNextPointForm(nextEntry, activeOperator?.name || '');

    setEntries((current) => {
      if (editingEntryId) {
        return current.map((entry) => (
          entry.__id === editingEntryId
            ? {
                ...entry,
                ...nextEntry,
                __removing: false,
              }
            : entry
        ));
      }

      return [{ ...nextEntry, __id: crypto.randomUUID() }, ...current];
    });
    setForm(nextPointSeed);
    setEditingEntryId(null);
    setNextPointReady(true);
    setConfirmedPosition(null);
    setDraftPosition(null);
    setGpsAccuracy(null);
    setManualCoordinateValue('');
    setManualCoordinateError('');
    if (markerRef.current) {
      markerRef.current.remove();
      markerRef.current = null;
    }
    setStep('location');
    setGeoStatus('Ponto salvo na fila local.');
    setToast(
      isEntryReadyToSync(form)
        ? (editingEntryId ? 'Ponto atualizado. Toque em Novo ponto para seguir.' : 'Ponto salvo. Toque em Novo ponto para seguir.')
        : 'Ponto salvo, mas ainda não está pronto para envio.'
    );
  }, [activeOperator, confirmedPosition, editingEntryId, entries, form]);

  const handleExport = useCallback(() => {
    if (!operatorEntries.length) {
      setToast('Adicione pelo menos um registro antes de exportar.');
      return;
    }
    exportWorkbook(operatorEntries);
    setToast('Planilha exportada no formato ideal para importação.');
  }, [operatorEntries]);

  const handleManagerExport = useCallback(async (scope) => {
    if (!activeOperator?.can_export) return;

    try {
      setManagerLoading(true);
      const rows = await fetchSubmissionExport(scope);
      if (!rows.length) {
        setToast(scope === 'pending' ? 'Não há registros pendentes para exportar.' : 'Não há registros para exportar.');
        return;
      }

      exportWorkbook(rows);

      if (scope === 'pending') {
        await markSubmissionsExported(rows.map((row) => row.client_uuid), activeOperator.id);
        await loadManagerSummary();
      }

      setToast(
        scope === 'pending'
          ? `${rows.length} registro(s) pendente(s) exportado(s).`
          : `${rows.length} registro(s) exportado(s) do histórico completo.`
      );
    } catch (error) {
      setToast(error.message || 'Não foi possível concluir a exportação.');
    } finally {
      setManagerLoading(false);
    }
  }, [activeOperator, loadManagerSummary]);

  const handleSyncSingleEntry = useCallback(async (entry) => {
    if (!isEntryReadyToSync(entry)) {
      setToast('Confirme o local e marque implantação concluída = SIM antes de enviar.');
      return;
    }

    await performSyncEntries([entry]);
  }, [performSyncEntries]);

  const handleRemoveEntry = useCallback((id) => {
    setEntries((current) => current.filter((entry) => entry.__id !== id));
    setEditingEntryId((current) => (current === id ? null : current));
  }, []);

  if (!activeOperator) {
    return (
      <div className="field-app auth-app">
        <section className="auth-card">
          <div className="hero-badge">Luz de Campo</div>
          <h1>Acesso de equipe</h1>

          <form className="auth-form" onSubmit={handleLogin}>
            <label className="form-field full">
              <span>Usuário autorizado</span>
              <select
                value={accessForm.operatorId}
                disabled={authLoading}
                onChange={(event) => handleAccessFieldChange('operatorId', event.target.value)}
              >
                <option value="">
                  {authLoading ? 'Carregando usuários...' : 'Selecione o usuário'}
                </option>
                {allowedOperators.map((operator) => (
                  <option key={operator.id} value={operator.id}>
                    {operator.name}
                  </option>
                ))}
              </select>
            </label>

            <label className="form-field full">
              <span>Codigo de acesso</span>
              <input
                type="password"
                value={accessForm.accessCode}
                placeholder="Digite o codigo de acesso"
                onChange={(event) => handleAccessFieldChange('accessCode', event.target.value)}
              />
            </label>

            {online && (
              <div className="form-field full">
                <span>RPAs para uso offline</span>
                <div className="offline-rpa-picker">
                  {rpaOptions.map((rpa) => {
                    const isSelected = accessForm.offlineRpas.includes(rpa);
                    return (
                      <button
                        key={rpa}
                        type="button"
                        className={`offline-rpa-chip${isSelected ? ' is-selected' : ''}`}
                        disabled={authLoading || offlinePrepLoading}
                        onClick={() => handleToggleOfflineRpa(rpa)}
                      >
                        {rpa}
                      </button>
                    );
                  })}
                </div>
                <div className="auth-help">
                  {accessForm.offlineRpas.length
                    ? `${accessForm.offlineRpas.length}/${MAX_OFFLINE_RPAS} RPAs selecionadas`
                    : `Selecione até ${MAX_OFFLINE_RPAS} RPAs`}
                </div>
              </div>
            )}

            {authError && <div className="auth-error">{authError}</div>}
            {authInfo && <div className="auth-help">{authInfo}</div>}
            <div className="auth-help">
              {online
                ? `Se for trabalhar sem sinal, escolha até ${MAX_OFFLINE_RPAS} RPAs antes de entrar. O app vai preparar o acesso e o mapa dessas areas.`
                : 'Sem internet: entram apenas acessos e areas offline ja preparados neste aparelho.'}
            </div>
            <button className="primary-action auth-submit" type="submit" disabled={offlinePrepLoading}>
              {offlinePrepLoading ? 'Preparando acesso offline...' : 'Entrar no sistema'}
            </button>
          </form>
        </section>
      </div>
    );
  }

  if (!activeOperator.can_export && !activeModule) {
    return (
      <div className="field-app module-app">
        <section className="module-shell">
          <div>
            <div className="hero-badge">Luz de Campo</div>
            <h1>Escolha o módulo</h1>
            <p>Selecione o tipo de registro que será feito em campo.</p>
          </div>
          <div className="module-operator-row">
            <span className="operator-pill">{activeOperator.name}</span>
            <button type="button" className="ghost-action" onClick={handleLogout}>
              Sair
            </button>
          </div>
          <div className="module-grid">
            {fieldModules.map((module) => (
              <button
                key={module.slug}
                type="button"
                className="module-card"
                onClick={() => handleSelectModule(module.slug)}
              >
                <span>{module.slug === 'sao-joao' ? 'Evento' : 'Operação'}</span>
                <strong>{module.name}</strong>
                <small>{module.description}</small>
              </button>
            ))}
          </div>
        </section>
        {toast && <div className="toast">{toast}</div>}
      </div>
    );
  }

  if (activeOperator.can_export) {
    return (
      <div className="field-app manager-app">
        <header className="hero">
          <div className="hero-badge">Luz de Campo</div>
            <h1>Painel gerencial</h1>
          <div className="hero-toolbar">
            <span className="operator-pill">{activeOperator.name}</span>
            <button type="button" className="hero-logout" onClick={handleLogout}>
              Sair
            </button>
          </div>
        </header>

        <main className="manager-layout">
          <section className="panel manager-panel manager-summary-panel">
            <div className="panel-header">
              <span className="panel-step">Gestão</span>
              <strong>Visão dos registros sincronizados</strong>
              <small>Baixe apenas os pendentes ou gere uma planilha completa do histórico.</small>
            </div>

            <div className="manager-stats">
              <article className="manager-stat-card">
                <strong>{managerSummary.total}</strong>
                <span>Total no banco</span>
              </article>
              <article className="manager-stat-card">
                <strong>{managerSummary.pending_export}</strong>
                <span>Ainda não exportados</span>
              </article>
              <article className="manager-stat-card">
                <strong>{managerSummary.exported}</strong>
                <span>Já exportados</span>
              </article>
            </div>

            <div className="manager-actions">
              <button
                type="button"
                className="primary-action"
                onClick={() => handleManagerExport('pending')}
                disabled={managerLoading || managerSummary.pending_export === 0}
              >
                {managerLoading ? 'Processando...' : 'Baixar pendentes'}
              </button>
              <button
                type="button"
                className="ghost-action"
                onClick={() => handleManagerExport('all')}
                disabled={managerLoading || managerSummary.total === 0}
              >
                Baixar histórico completo
              </button>
            </div>

            <div className="manager-refresh">
              <button
                type="button"
                className="shortcut-action shortcut-action-light"
                onClick={() => {
                  loadManagerSummary();
                  loadManagerRows();
                }}
                disabled={managerLoading}
              >
                Atualizar contadores
              </button>
            </div>

            <div className="manager-modules">
              <div className="manager-modules-header">
                <strong>Módulos de campo</strong>
                <span>Ative apenas o que a equipe deve usar.</span>
              </div>
              {fieldModules.map((module) => (
                <article key={module.slug} className="manager-module-card">
                  <div>
                    <strong>{module.name}</strong>
                    <span>{module.description}</span>
                  </div>
                  <button
                    type="button"
                    className={`manager-module-toggle${module.active ? ' is-active' : ''}`}
                    onClick={() => handleToggleManagerModule(module)}
                    disabled={managerModuleActionSlug === module.slug}
                  >
                    {managerModuleActionSlug === module.slug ? 'Salvando...' : module.active ? 'Ativo' : 'Inativo'}
                  </button>
                </article>
              ))}
            </div>
          </section>

          <section className={`panel manager-panel manager-table-panel${managerTableMaximized ? ' is-maximized' : ''}`}>
            <div className="panel-header">
              <div>
                <span className="panel-step">Registros</span>
                <strong>Pontos registrados</strong>
                <small>{managerTableRows.length} linha(s) no recorte ativo</small>
              </div>
              <button
                type="button"
                className="manager-maximize-button"
                onClick={() => setManagerTableMaximized((current) => !current)}
              >
                {managerTableMaximized ? 'Voltar' : 'Maximizar'}
              </button>
            </div>

            {managerRowsLoading ? (
              <div className="placeholder-card">Carregando pontos registrados...</div>
            ) : managerRows.length === 0 ? (
              <div className="placeholder-card">Nenhum ponto registrado no banco.</div>
            ) : (
              <div className="manager-table-scroll">
                <table className="manager-submission-table">
                  <thead>
                    <tr>
                      <th>Item</th>
                      {MANAGER_COMPACT_TABLE_FIELDS.map((field) => (
                        <th key={field}>{MANAGER_FIELD_LABELS[field] || field}</th>
                      ))}
                      <th>Detalhes</th>
                      <th className="manager-table-actions-col">Ações</th>
                    </tr>
                    <tr className="manager-table-filter-row">
                      <th />
                      {MANAGER_COMPACT_TABLE_FIELDS.map((field) => (
                        <th key={`filter-${field}`}>
                          <input
                            value={managerTableFilters[field] || ''}
                            placeholder="filtrar..."
                            onChange={(event) => handleManagerFilterChange(field, event.target.value)}
                          />
                        </th>
                      ))}
                      <th />
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {managerTableRows.map(({ row, exportRow, clientUuid, index }) => {
                      const rowClientUuid = clientUuid;
                      const isEditing = managerEditingClientUuid === rowClientUuid;
                      return (
                        <tr key={rowClientUuid || `${row.operador || 'ponto'}-${row.synced_em || ''}`}>
                          <td data-label="Item" className="manager-table-item-cell">{index + 1}</td>
                          {MANAGER_COMPACT_TABLE_FIELDS.map((field) => (
                            <td key={field} data-label={MANAGER_FIELD_LABELS[field] || field}>
                              {renderManagerTableCell(field, exportRow, isEditing)}
                            </td>
                          ))}
                          <td data-label="Detalhes" className="manager-table-details-cell">
                            <details className="manager-row-details">
                              <summary>Ver dados</summary>
                              {renderManagerDetailGroup(MANAGER_DETAIL_ONLY_FIELDS, exportRow, isEditing)}
                            </details>
                          </td>
                          <td data-label="Ações" className="manager-table-actions-cell">
                            <div className="manager-table-actions">
                              {isEditing ? (
                                <>
                                  <button type="button" className="primary-action manager-table-button" onClick={() => handleSaveManagerEdit(rowClientUuid)} disabled={managerRowActionId === rowClientUuid}>
                                    Salvar
                                  </button>
                                  <button type="button" className="ghost-action manager-table-button" onClick={() => setManagerEditingClientUuid('')} disabled={managerRowActionId === rowClientUuid}>
                                    Cancelar
                                  </button>
                                </>
                              ) : (
                                <>
                                  <button type="button" className="ghost-action manager-table-button" onClick={() => handleStartManagerEdit(row)}>
                                    Editar
                                  </button>
                                  <button type="button" className="ghost-action manager-table-button manager-table-delete" onClick={() => handleDeleteManagerRow(row)} disabled={managerRowActionId === rowClientUuid || !rowClientUuid}>
                                    Excluir
                                  </button>
                                </>
                              )}
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </section>

          <section className="panel manager-panel manager-users-panel">
            <div className="panel-header">
              <span className="panel-step">Acessos</span>
              <strong>Cadastrar novo usuário</strong>
              <small>O João pode liberar rapidamente quem vai operar em campo e definir se a pessoa também exporta.</small>
            </div>

            <form className="manager-user-form" onSubmit={handleCreateManagerUser}>
              <label className="form-field full">
                <span>Código gerencial</span>
                <input
                  type="password"
                  value={managerUserForm.managerAccessCode}
                    placeholder="Digite o código gerencial"
                  onChange={(event) => handleManagerUserFieldChange('managerAccessCode', event.target.value)}
                />
              </label>

              <label className="form-field full">
                <span>Nome do usuário</span>
                <input
                  type="text"
                  value={managerUserForm.name}
                  placeholder="Ex.: Maria da Silva"
                  onChange={(event) => handleManagerUserFieldChange('name', event.target.value)}
                />
              </label>

              <label className="form-field">
                <span>Código de acesso</span>
                <input
                  type="text"
                  value={managerUserForm.accessCode}
                  placeholder="Ex.: 4821"
                  onChange={(event) => handleManagerUserFieldChange('accessCode', event.target.value)}
                />
              </label>

              <label className="manager-user-toggle">
                <input
                  type="checkbox"
                  checked={managerUserForm.canExport}
                  onChange={(event) => handleManagerUserFieldChange('canExport', event.target.checked)}
                />
                <div>
                  <strong>Pode exportar planilhas</strong>
                  <span>Ative apenas para perfil gerencial.</span>
                </div>
              </label>

              <button type="submit" className="primary-action" disabled={managerUserLoading}>
                {managerUserLoading ? 'Salvando...' : 'Adicionar usuário'}
              </button>
            </form>

            <div className="manager-user-list-header">
              <div>
                <strong>Usuarios liberados</strong>
                <span>{allowedOperators.length} acesso(s) ativo(s)</span>
              </div>
              <button
                type="button"
                className="manager-user-collapse"
                onClick={() => setManagerUsersCollapsed((current) => !current)}
              >
                {managerUsersCollapsed ? 'Mostrar' : 'Ocultar'}
              </button>
            </div>

            {!managerUsersCollapsed && (
              <div className="manager-user-list">
              {allowedOperators.map((operator) => (
                <article key={operator.id} className="manager-user-card">
                  <div className="manager-user-card-copy">
                    <strong>{operator.name}</strong>
                    <span>{operator.can_export ? 'Perfil gerencial' : 'Equipe de campo'}</span>
                  </div>
                  <div className="manager-user-card-actions">
                    {operator.id === activeOperator.id ? (
                      <span className="manager-user-card-self">Usuário atual</span>
                    ) : (
                      <button
                        type="button"
                        className="ghost-action manager-user-remove"
                        onClick={() => handleDeactivateManagerUser(operator)}
                        disabled={managerUserActionId === operator.id}
                      >
                        {managerUserActionId === operator.id ? 'Desativando...' : 'Desativar'}
                      </button>
                    )}
                  </div>
                </article>
              ))}
              </div>
            )}
          </section>
        </main>

        {toast && <div className="toast">{toast}</div>}
      </div>
    );
  }

  if (activeModule === 'sao-joao') {
    const pendingSaoJoaoEntries = saoJoaoEntries.filter((entry) => entry.__syncStatus !== 'synced');

    return (
      <div className="field-app sao-joao-app">
        <header className="hero sao-joao-hero">
          <div className="hero-badge">Luz de Campo</div>
          <h1>São João</h1>
          <div className="hero-toolbar">
            <span className="operator-pill">{activeOperator.name}</span>
            <div className="module-header-actions">
              <button type="button" className="hero-logout" onClick={handleBackToModules}>
                Trocar módulo
              </button>
              <button type="button" className="hero-logout" onClick={handleLogout}>
                Sair
              </button>
            </div>
          </div>
        </header>

        <main className="sao-joao-layout">
          <section className="panel sao-joao-form-panel">
            <div className="panel-header">
              <span className="panel-step">Checklist</span>
              <strong>Vistoria do polo</strong>
              <small>Preencha somente o que for verificado. As fotos são opcionais e ficam vinculadas à pergunta.</small>
            </div>

            <div className="sao-joao-top-grid">
              <label className="form-field full">
                <span>Nome do polo</span>
                <input
                  value={saoJoaoForm.POLO_NOME}
                  placeholder="Ex.: Polo Casa Amarela"
                  onChange={(event) => handleSaoJoaoFieldChange('POLO_NOME', event.target.value)}
                />
              </label>

              <label className="form-field">
                <span>Latitude</span>
                <input
                  value={saoJoaoForm.LATITUDE}
                  placeholder="-8.052240"
                  onChange={(event) => handleSaoJoaoFieldChange('LATITUDE', event.target.value)}
                />
              </label>

              <label className="form-field">
                <span>Longitude</span>
                <input
                  value={saoJoaoForm.LONGITUDE}
                  placeholder="-34.928610"
                  onChange={(event) => handleSaoJoaoFieldChange('LONGITUDE', event.target.value)}
                />
              </label>

              <div className="sao-joao-location-actions">
                <button type="button" className="primary-action" onClick={handleCaptureSaoJoaoLocation}>
                  Usar GPS
                </button>
                <input
                  type="text"
                  placeholder="Cole lat, long e pressione Enter"
                  onKeyDown={(event) => {
                    if (event.key === 'Enter') {
                      event.preventDefault();
                      handlePasteSaoJoaoCoordinates(event.currentTarget.value);
                      event.currentTarget.value = '';
                    }
                  }}
                  onPaste={(event) => {
                    window.setTimeout(() => {
                      handlePasteSaoJoaoCoordinates(event.currentTarget.value);
                      event.currentTarget.value = '';
                    }, 0);
                  }}
                />
              </div>
            </div>

            <div className="sao-joao-question-groups">
              {SAO_JOAO_QUESTIONS.map((group) => (
                <section key={group.group} className="sao-joao-question-group">
                  <h2>{group.group}</h2>
                  {group.items.map((question) => (
                    <article key={question.key} className="sao-joao-question-card">
                      <div className="sao-joao-question-copy">
                        <strong>{question.label}</strong>
                      </div>
                      <div className="quick-picks">
                        {SAO_JOAO_ANSWER_OPTIONS.map((option) => (
                          <button
                            key={option}
                            type="button"
                            className={`quick-pick${saoJoaoForm.RESPOSTAS?.[question.key] === option ? ' active' : ''}`}
                            onClick={() => handleSaoJoaoAnswerChange(question.key, option)}
                          >
                            {option}
                          </button>
                        ))}
                      </div>

                      {(question.photos || []).map((photo) => {
                        const image = saoJoaoForm.IMAGENS?.[photo.key];
                        return (
                          <div key={photo.key} className="sao-joao-photo-slot">
                            <span>{photo.label}</span>
                            <div className="image-actions">
                              <label className="image-upload-action">
                                <input
                                  type="file"
                                  accept="image/*"
                                  capture="environment"
                                  className="image-upload-input"
                                  onChange={(event) => handleSaoJoaoImageSelection(photo.key, event)}
                                />
                                {saoJoaoImageLoadingKey === photo.key ? 'Preparando...' : image ? 'Trocar foto' : 'Adicionar foto'}
                              </label>
                              {image && (
                                <button type="button" className="shortcut-action shortcut-action-light" onClick={() => handleRemoveSaoJoaoImage(photo.key)}>
                                  Remover
                                </button>
                              )}
                            </div>
                            {image && (
                              <div className="image-preview-card sao-joao-preview-card">
                                <img src={image.data_url} alt={photo.label} className="image-preview" />
                                <div className="image-preview-copy">
                                  <strong>{image.filename}</strong>
                                  <span>{Math.round(image.size_bytes / 1024)} KB</span>
                                </div>
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </article>
                  ))}
                </section>
              ))}
            </div>

            <label className="form-field full">
              <span>Observações</span>
              <textarea
                className="sao-joao-textarea"
                value={saoJoaoForm.OBSERVACOES}
                placeholder="Registre algo importante para a equipe."
                onChange={(event) => handleSaoJoaoFieldChange('OBSERVACOES', event.target.value)}
              />
            </label>

            <button type="button" className="save-action" onClick={handleSaveSaoJoaoEntry}>
              Salvar checklist
            </button>
          </section>

          <aside className="panel sao-joao-queue-panel">
            <div className="panel-header">
              <span className="panel-step">Fila local</span>
              <strong>{pendingSaoJoaoEntries.length} pendente(s)</strong>
              <small>Os checklists ficam no aparelho até sincronizar.</small>
            </div>
            <button
              type="button"
              className="primary-action"
              onClick={handleSyncSaoJoaoEntries}
              disabled={!pendingSaoJoaoEntries.length || saoJoaoSyncing || !online}
            >
              {saoJoaoSyncing ? 'Sincronizando...' : 'Sincronizar São João'}
            </button>

            <div className="sao-joao-entry-list">
              {!pendingSaoJoaoEntries.length ? (
                <div className="placeholder-card">Nenhum checklist salvo.</div>
              ) : pendingSaoJoaoEntries.map((entry) => (
                <article key={entry.__id} className="queue-item sao-joao-entry-card">
                  <div>
                    <strong>{entry.POLO_NOME || 'Polo sem nome'}</strong>
                    <span>{entry.OPERADOR || activeOperator.name}</span>
                    <small>{entry.LATITUDE || '-'}, {entry.LONGITUDE || '-'}</small>
                  </div>
                  <button type="button" className="ghost-action" onClick={() => handleRemoveSaoJoaoEntry(entry.__id)}>
                    Remover
                  </button>
                </article>
              ))}
            </div>
          </aside>
        </main>

        {toast && <div className="toast">{toast}</div>}
      </div>
    );
  }

  return (
    <div className="field-app">
      <header className="hero">
        <div className="hero-badge">Luz de Campo</div>
        <h1>Captura rápida para equipe de campo</h1>
        <div className="hero-toolbar">
          <span className="operator-pill">
            {activeOperator.name}
            {activeOperator.can_export ? ' · Exportador' : ''}
          </span>
          <div className="module-header-actions">
            <button type="button" className="hero-logout" onClick={handleBackToModules}>
              Trocar módulo
            </button>
            <button type="button" className="hero-logout" onClick={handleLogout}>
              Sair
            </button>
          </div>
        </div>
      </header>

      <main className="field-layout">
        <section className="panel panel-location">
          <div className={`map-stage${needsLocationConfirmation ? ' map-stage-confirming' : ''}`}>
            <div ref={mapContainerRef} className="map-shell" />

            <div className="map-gps-tools">
              <button
                className="map-gps-recenter"
                type="button"
                onClick={handleRecenterGps}
                aria-label="Calibrar GPS e recentralizar pin"
                title="Calibrar GPS"
              >
                GPS
              </button>
              {gpsAccuracyLabel ? (
                <div className={`map-gps-accuracy map-gps-accuracy-${gpsAccuracyTone}`}>
                  Precisao {gpsAccuracyLabel}
                </div>
              ) : null}
            </div>

            <div className="map-bottomsheet">
              <form className="manual-coordinate-card" onSubmit={handleManualCoordinateSubmit}>
                <div className="manual-coordinate-copy">
                  <strong>Colar coordenadas</strong>
                  <span>Cole latitude e longitude com Ctrl+V.</span>
                </div>
                <div className="manual-coordinate-row">
                  <input
                    type="text"
                    className="manual-coordinate-input"
                    value={manualCoordinateValue}
                    placeholder="-8.052240, -34.928610"
                    onChange={(event) => {
                      setManualCoordinateValue(event.target.value);
                      if (manualCoordinateError) setManualCoordinateError('');
                    }}
                    onPaste={handleManualCoordinatePaste}
                    aria-label="Colar latitude e longitude"
                  />
                  <button type="submit" className="ghost-action manual-coordinate-apply">
                    Usar
                  </button>
                </div>
                {manualCoordinateError && <div className="manual-coordinate-error">{manualCoordinateError}</div>}
              </form>
              <div className="location-actions">
                <button
                  className={`capture-action${needsLocationConfirmation ? ' capture-action-confirm' : ''}`}
                  type="button"
                  onClick={handleLocationAction}
                >
                  {locationActionLabel}
                </button>
              </div>
            </div>
          </div>

        </section>

        <section ref={formSectionRef} className="panel panel-form">
          <div className="panel-header">
            <span className="panel-step">Etapa 2</span>
            <strong>Preencher o ponto</strong>
            <small>Depois de confirmar o local, preencha os dados.</small>
          </div>

          {step !== 'form' ? (
            <div className="placeholder-card">
              Primeiro confirme o local no mapa.
            </div>
          ) : (
            <div className="form-grid">
              <div className="form-shortcuts">
                <button
                  type="button"
                  className="shortcut-action"
                  onClick={handleReuseLastEntry}
                  disabled={!lastEntryTemplate}
                >
                  Repetir último padrão
                </button>
                <button
                  type="button"
                  className="shortcut-action shortcut-action-light"
                  onClick={handleClearOptionalFields}
                >
                  Limpar opcionais
                </button>
              </div>

                <div className="form-status-card" aria-label="Progresso do formulário">
                  <div className="form-progress">
                    <div className="form-progress-copy">
                      <strong>{requiredFieldsFilled} de {totalRequiredFields}</strong>
                      <span>{remainingRequiredFields === 0 ? 'Tudo preenchido' : `${remainingRequiredFields} campo(s) ainda opcionais`}</span>
                    </div>
                    <div className="form-progress-track" aria-hidden="true">
                      <div className="form-progress-bar" style={{ width: `${formProgressPercent}%` }} />
                    </div>
                  </div>
                </div>

              {FORM_FIELDS.map((field) => {
                if (field === 'OBRA_NOME' && form.MOTIVO_IMPLANTACAO !== 'OBRA') {
                  return null;
                }
                const isCoordinates = field === 'LATITUDE' || field === 'LONGITUDE';
                const isQuantityField = field === 'QTDE';
                const inputType = ['QTDE', 'PERDAS', 'TOTAL_CARGA', 'CONSUMO_kW', 'CONSUMO_kW_MES'].includes(field)
                  ? 'number'
                  : field === 'ATUALIZACAO'
                    ? 'date'
                    : 'text';
                const hasValue = String(form[field] ?? '').trim().length > 0;

                return (
                  <label
                    key={field}
                    className={`form-field${field === 'ENDERECO' ? ' full' : ''}${hasValue ? ' is-filled' : ' is-empty'}${guidedFieldKey === field ? ' is-guided' : ''}`}
                  >
                    <span>
                      {FIELD_LABELS[field]}
                    </span>
                    {!isQuantityField && (
                      <input
                        ref={(element) => {
                          if (element) {
                            formFieldRefs.current[field] = element;
                          } else {
                            delete formFieldRefs.current[field];
                          }
                        }}
                        type={inputType}
                        inputMode={inputType === 'number' ? 'decimal' : undefined}
                        readOnly={isCoordinates}
                        value={form[field]}
                        placeholder={FIELD_PLACEHOLDERS[field] || ''}
                        onChange={(event) => handleChangeField(field, event.target.value)}
                        onKeyDown={(event) => handleFieldKeyDown(field, event)}
                      />
                    )}
                    {QUICK_OPTIONS[field] && (
                      <div className="quick-picks">
                        {QUICK_OPTIONS[field].map((option) => {
                          const optionValue = typeof option === 'string' ? option : option.value;
                          const optionLabel = typeof option === 'string' ? option : option.label;

                          return (
                            <button
                              key={optionValue}
                              type="button"
                              className={`quick-pick${form[field] === optionValue ? ' active' : ''}`}
                              onClick={() => handleApplyQuickOption(field, optionValue)}
                            >
                              {optionLabel}
                            </button>
                          );
                        })}
                      </div>
                    )}
                  </label>
                );
              })}

              <div className="luminaire-section">
                {(form.LUMINARIAS || []).map((item, index) => {
                  const hasPotencia = String(item?.POTENCIA ?? '').trim().length > 0;
                  const hasImagem = Boolean(item?.IMAGEM);
                  return (
                    <div key={item.INDICE} className="form-image-card luminaire-card">
                      <div className="panel-header panel-header-inline">
                        <span className="panel-step">Luminária {item.INDICE}</span>
                        <strong>Potência e foto</strong>
                        <small>Preencha a potência e adicione a foto desta luminária.</small>
                      </div>
                      <label className={`form-field full${hasPotencia ? ' is-filled' : ' is-empty'}${guidedFieldKey === `LUMINARIA_${item.INDICE}_POTENCIA` ? ' is-guided' : ''}`}>
                        <span>Potência da luminária {item.INDICE}</span>
                        <input
                          ref={(element) => {
                            const focusKey = `LUMINARIA_${item.INDICE}_POTENCIA`;
                            if (element) {
                              formFieldRefs.current[focusKey] = element;
                            } else {
                              delete formFieldRefs.current[focusKey];
                            }
                          }}
                          type="number"
                          inputMode="decimal"
                          value={item.POTENCIA}
                          placeholder="Ex.: 70"
                          onChange={(event) => handleLuminairePotenciaChange(index, event.target.value)}
                          onKeyDown={(event) => handleLuminaireFieldKeyDown(`LUMINARIA_${item.INDICE}_POTENCIA`, event)}
                        />
                      </label>
                      <div className={`form-field full${hasImagem ? ' is-filled' : ' is-empty'}`}>
                        <span>Foto da luminária {item.INDICE}</span>
                      </div>
                      <div className="image-actions">
                        <input
                          ref={(element) => {
                            if (element) {
                              luminaireFileInputRefs.current[item.INDICE] = element;
                            } else {
                              delete luminaireFileInputRefs.current[item.INDICE];
                            }
                          }}
                          type="file"
                          accept="image/*"
                          capture="environment"
                          className="image-upload-input"
                          onChange={(event) => handleImageSelection(index, event)}
                        />
                        <button
                          ref={(element) => {
                            const focusKey = `LUMINARIA_${item.INDICE}_IMAGEM`;
                            if (element) {
                              formFieldRefs.current[focusKey] = element;
                            } else {
                              delete formFieldRefs.current[focusKey];
                            }
                          }}
                          type="button"
                          className={`image-upload-action${guidedFieldKey === `LUMINARIA_${item.INDICE}_IMAGEM` ? ' is-guided' : ''}`}
                          onClick={() => luminaireFileInputRefs.current[item.INDICE]?.click()}
                        >
                          {item.IMAGEM ? 'Trocar foto' : 'Adicionar foto'}
                        </button>
                        {item.IMAGEM && (
                          <button type="button" className="shortcut-action shortcut-action-light" onClick={() => handleRemoveImage(index)}>
                            Remover foto
                          </button>
                        )}
                      </div>
                      {item.IMAGEM && (
                        <div className="image-preview-card">
                          <img src={item.IMAGEM.data_url} alt={`Prévia da luminária ${item.INDICE}`} className="image-preview" />
                          <div className="image-preview-copy">
                            <strong>{item.IMAGEM.filename}</strong>
                            <span>{Math.round(item.IMAGEM.size_bytes / 1024)} KB</span>
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>

              <button ref={saveActionRef} className={`save-action${guidedFieldKey === 'SAVE_ACTION' ? ' is-guided' : ''}`} type="button" onClick={handleSaveEntry}>
                {editingEntryId ? 'Atualizar registro local' : 'Adicionar registro'}
              </button>
            </div>
          )}
        </section>
      </main>

      {step === 'location' && operatorEntries.length > 0 && (
        <div className="next-point-slot">
          <button
            type="button"
            className={`next-point-action${nextPointReady ? ' is-highlighted' : ''}`}
            onClick={handleStartNextPoint}
          >
            Novo ponto
          </button>
        </div>
      )}

        <section className={`panel queue-panel${syncing ? ' queue-panel-syncing' : ''}`}>
          <div className="panel-header">
            <span className="panel-step">Etapa 3</span>
            <strong>Fila local</strong>
            <small>Os pontos ficam no aparelho e sobem para o banco quando houver conexão.</small>
          </div>

        {syncing && (
          <div className="queue-sync-banner" role="status" aria-live="polite">
            <span className="queue-sync-spinner" aria-hidden="true" />
            <div className="queue-sync-copy">
              <strong>Sincronizando com o banco</strong>
              <span>Enviando os pontos prontos e limpando a fila local.</span>
            </div>
          </div>
        )}

        <div className="queue-toolbar">
          <div className="queue-toolbar-main">
            <div className="queue-counter">
              <strong>{operatorEntries.length}</strong>
              <span>
                {syncablePendingEntries.length} pronto(s) para envio · {awaitingConfirmationEntries.length} pendente(s) de confirmação · {online ? 'online' : 'offline'}
              </span>
            </div>
            <div className="queue-filters" role="tablist" aria-label="Filtrar fila local">
              <button
                type="button"
                className={`queue-filter${queueFilter === 'all' ? ' active' : ''}`}
                onClick={() => setQueueFilter('all')}
              >
                Todos
              </button>
              <button
                type="button"
                className={`queue-filter${queueFilter === 'ready' ? ' active' : ''}`}
                onClick={() => setQueueFilter('ready')}
              >
                Prontos para envio
              </button>
              <button
                type="button"
                className={`queue-filter${queueFilter === 'waiting' ? ' active' : ''}`}
                onClick={() => setQueueFilter('waiting')}
              >
                Pendentes
              </button>
            </div>
          </div>
          <div className="queue-actions">
            {activeOperator.can_export && (
              <button className="ghost-action queue-export-action" type="button" onClick={handleExport}>
                Exportar planilha
              </button>
            )}
            <button className="primary-action" type="button" onClick={handleSyncEntries} disabled={!syncablePendingEntries.length || syncing || !online}>
              {syncing ? 'Sincronizando...' : 'Sincronizar pontos'}
            </button>
          </div>
        </div>

          {!operatorEntries.length ? (
            <div className="placeholder-card">
              Nenhum ponto salvo ainda. Capture o local e adicione o primeiro ponto.
            </div>
        ) : !filteredOperatorEntries.length ? (
          <div className="placeholder-card">
            Nenhum ponto encontrado para esse filtro agora.
          </div>
        ) : (
          <div className="queue-list">
            {filteredOperatorEntries.map((entry) => (
              <article
                key={entry.__id}
                className={`queue-item${entry.__syncStatus === 'synced' ? ' queue-item-synced' : ''}${entry.__removing ? ' queue-item-removing' : ''}`}
              >
                <div>
                  <strong>{entry.ENDERECO || 'Sem endereço'}</strong>
                  <span>{entry.BAIRRO || 'Sem bairro'} · RPA {entry.RPA || '-'}</span>
                  <span>{entry.TIPO_IMPLANTACAO || 'Situação não informada'} · Implantação {entry.IMPLANTACAO_CONCLUIDA === 'SIM' ? 'concluída' : 'em aberto'}</span>
                  {countLuminaireImages(entry.LUMINARIAS || []) > 0 && (
                    <span>{countLuminaireImages(entry.LUMINARIAS || [])} foto(s) anexada(s)</span>
                  )}
                  <small>
                    {entry.LATITUDE}, {entry.LONGITUDE} · {entry.__syncStatus === 'synced' ? 'Sincronizado' : isEntryReadyToSync(entry) ? 'Pronto para envio' : 'Pendente de confirmação'}
                  </small>
                </div>
                <div className="queue-item-actions">
                  {!isEntryReadyToSync(entry) && (
                    <span className="queue-warning">Confirme local e conclusão</span>
                  )}
                  <button type="button" className="shortcut-action shortcut-action-light" onClick={() => handleEditEntry(entry)}>
                    Editar
                  </button>
                  <button
                    type="button"
                    className="primary-action queue-sync-action"
                    onClick={() => handleSyncSingleEntry(entry)}
                    disabled={!online || syncing || !isEntryReadyToSync(entry)}
                  >
                    Enviar
                  </button>
                  <button type="button" className="ghost-action" onClick={() => handleRemoveEntry(entry.__id)}>
                    Remover
                  </button>
                </div>
              </article>
            ))}
          </div>
        )}
      </section>

      {toast && <div className="toast">{toast}</div>}
    </div>
  );
}
