// Lógica do app: formulário, fotos, histórico (IndexedDB) e exportação para Word.
// Sem login, sem conta externa, sem nuvem — tudo roda e fica salvo no aparelho.

const TYPES = {
  interna: { label: 'Inspeção Interna', color: '#4C7A5D', tint: '#EAF2EC', last: 'Última inspeção interna', next: 'Próxima Inspeção Interna', sec: 'Inspeção Interna' },
  externa: { label: 'Inspeção Externa', color: '#C6821F', tint: '#FBF0DC', last: 'Última inspeção externa', next: 'Próxima Inspeção Externa', sec: 'Inspeção Externa' },
  tubulacao: { label: 'Inspeção de Tubulação', color: '#3E6E8E', tint: '#E9F1F5', last: 'Última inspeção externa', next: 'Inspeção Externa (referência)', sec: 'Inspeção de Tubulação' }
};

let currentType = 'interna';
let photos = [];
const itemPhotos = {};
let currentRecordId = null;
let isoSeq = 0;

// ---------------------------------------------------------------------
// Navegação entre vistas (Formulário / Histórico)
// ---------------------------------------------------------------------

function showFormView() {
  document.getElementById('formView').classList.remove('hidden');
  document.getElementById('historyView').classList.add('hidden');
  document.getElementById('formBar').classList.remove('hidden');
  document.getElementById('navlist').classList.remove('hidden');
  document.getElementById('navFormBtn').classList.add('active');
  document.getElementById('navHistoryBtn').classList.remove('active');
}

function showHistoryView() {
  document.getElementById('formView').classList.add('hidden');
  document.getElementById('historyView').classList.remove('hidden');
  document.getElementById('formBar').classList.add('hidden');
  document.getElementById('navlist').classList.add('hidden');
  document.getElementById('navFormBtn').classList.remove('active');
  document.getElementById('navHistoryBtn').classList.add('active');
  renderHistoryView();
}

document.getElementById('navFormBtn').addEventListener('click', showFormView);
document.getElementById('navHistoryBtn').addEventListener('click', showHistoryView);
document.getElementById('newInspectionBtn').addEventListener('click', () => {
  startNewInspection();
  showFormView();
});

// ---------------------------------------------------------------------
// Tipo de inspeção / navegação lateral por seção
// ---------------------------------------------------------------------

function setType(type) {
  currentType = type;
  const t = TYPES[type];
  document.documentElement.style.setProperty('--current-color', t.color);
  document.documentElement.style.setProperty('--current-tint', t.tint);

  document.querySelectorAll('.type-btn').forEach(b => {
    b.dataset.active = (b.dataset.type === type) ? 'true' : 'false';
  });

  document.getElementById('plateBadge').textContent = t.label;
  document.getElementById('plateBadge').style.background = t.color;
  document.getElementById('lblUltima').textContent = t.last;
  document.getElementById('lblProxima').textContent = t.next;
  document.getElementById('secEspecificaTitle').textContent = t.sec;
  document.querySelector('[data-field="tipoInspecaoTxt"]').value = t.label;

  document.querySelectorAll('[data-section]').forEach(sec => {
    const types = (sec.dataset.types || '').split(',');
    sec.classList.toggle('hidden-type', !types.includes(type));
  });
  document.querySelectorAll('[data-types-only]').forEach(el => {
    const types = el.dataset.typesOnly.split(',');
    el.style.display = types.includes(type) ? '' : 'none';
  });

  renumberAndBuildNav();
  updateStatus();
}

function renumberAndBuildNav() {
  const navlist = document.getElementById('navlist');
  navlist.innerHTML = '';
  const visible = Array.from(document.querySelectorAll('[data-section]')).filter(s => !s.classList.contains('hidden-type'));
  visible.forEach((sec, i) => {
    const num = String(i + 1).padStart(2, '0');
    sec.querySelector('.num').textContent = num;
    const title = sec.querySelector('h2').textContent.replace(/^\d+/, '').trim();
    const secId = 'sec-auto-' + i;
    sec.id = secId;
    const li = document.createElement('li');
    li.innerHTML = `<button data-target="${secId}"><span class="dot" data-dot="${secId}"></span>${title}</button>`;
    navlist.appendChild(li);
  });
  navlist.querySelectorAll('button').forEach(btn => {
    btn.addEventListener('click', () => {
      showFormView();
      document.getElementById(btn.dataset.target).scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  });
  observeSections();
  if (typeof relabelIsometricos === 'function') relabelIsometricos();
}

let io;
function observeSections() {
  if (io) io.disconnect();
  io = new IntersectionObserver((entries) => {
    entries.forEach(e => {
      if (e.isIntersecting) {
        document.querySelectorAll('.navlist button').forEach(b => b.classList.remove('active'));
        const btn = document.querySelector(`.navlist button[data-target="${e.target.id}"]`);
        if (btn) btn.classList.add('active');
      }
    });
  }, { rootMargin: "-20% 0px -70% 0px" });
  document.querySelectorAll('[data-section]:not(.hidden-type)').forEach(s => io.observe(s));
}

document.querySelectorAll('.type-btn').forEach(btn => {
  btn.addEventListener('click', () => setType(btn.dataset.type));
});

function syncPlate() {
  document.getElementById('plateUnidade').textContent = document.querySelector('[data-field="unidade"]').value || '—';
  document.getElementById('plateLocal').textContent = document.querySelector('[data-field="local"]').value || '—';
  document.getElementById('plateTag').textContent = document.querySelector('[data-field="tag"]').value || '— — —';
}

function updateStatus() {
  let filled = 0, total = 0;
  document.querySelectorAll('[data-section]:not(.hidden-type)').forEach(sec => {
    total++;
    const inputs = Array.from(sec.querySelectorAll('input[data-field], textarea[data-field]')).filter(el => el.offsetParent !== null);
    const hasContent = inputs.some(i => i.type === 'radio' ? i.checked : (i.value && i.value.trim() !== ""));
    const hasPhotos = sec.querySelector('#photoGrid') && photos.length > 0;
    const photoKeys = Array.from(sec.querySelectorAll('[data-photo-key]')).map(h => h.dataset.photoKey);
    const hasItemPhotos = photoKeys.some(k => (itemPhotos[k] || []).length > 0);
    const dot = document.querySelector(`[data-dot="${sec.id}"]`);
    if (hasContent || hasPhotos || hasItemPhotos) { filled++; if (dot) dot.classList.add('done'); }
    else if (dot) { dot.classList.remove('done'); }
  });
  document.getElementById('statusText').textContent = `${filled} de ${total} seções com conteúdo`;
}

document.body.addEventListener('input', (e) => {
  if (e.target.matches('input[data-field], textarea[data-field], select[data-field]')) updateStatus();
});

// ---------------------------------------------------------------------
// Fotos gerais
// ---------------------------------------------------------------------

const photoGrid = document.getElementById('photoGrid');
const addPhotoBtn = photoGrid.querySelector('.add-photo');

function addPhotos(evt) {
  const files = Array.from(evt.target.files || []);
  files.forEach(file => {
    const reader = new FileReader();
    reader.onload = () => {
      const id = 'p' + Date.now() + Math.random().toString(36).slice(2, 6);
      photos.push({ id, src: reader.result, caption: '' });
      renderPhotos(); updateStatus();
    };
    reader.readAsDataURL(file);
  });
  evt.target.value = '';
}

function renderPhotos() {
  photoGrid.querySelectorAll('.photo-card').forEach(n => n.remove());
  photos.forEach(p => {
    const card = document.createElement('div');
    card.className = 'photo-card';
    card.innerHTML = `
      <div class="thumb"><img src="${p.src}" alt=""></div>
      <div class="cap"><input type="text" placeholder="Legenda da foto…" value="${p.caption.replace(/"/g, '&quot;')}" data-photo-caption="${p.id}"></div>
      <button class="rm" data-photo-remove="${p.id}">Remover foto</button>`;
    photoGrid.insertBefore(card, addPhotoBtn);
  });
  photoGrid.querySelectorAll('[data-photo-caption]').forEach(inp => {
    inp.addEventListener('input', () => { const p = photos.find(x => x.id === inp.dataset.photoCaption); if (p) p.caption = inp.value; });
  });
  photoGrid.querySelectorAll('[data-photo-remove]').forEach(btn => {
    btn.addEventListener('click', () => { photos = photos.filter(x => x.id !== btn.dataset.photoRemove); renderPhotos(); updateStatus(); });
  });
}

// ---------------------------------------------------------------------
// Isométricos repetíveis (Inspeção de Tubulação)
// ---------------------------------------------------------------------

function isoCardHTML(uid) {
  return `<div class="iso-card" data-iso="${uid}">
    <div class="iso-head">
      <span class="iso-num-badge" data-iso-badge>—</span>
      <input type="text" placeholder='Identificação do isométrico / linha (ex.: ISO-014, 6"-P-1023-A1CS)' data-field="tIsoId_${uid}">
      <button type="button" class="iso-remove" data-iso-remove>Remover isométrico</button>
    </div>
    <div class="iso-body">
      <div class="yn">
        <span style="font-size:12.5px; font-weight:600; color:var(--steel); text-transform:uppercase;">Possui isolamento térmico?</span>
        <label><input type="radio" name="isolamento_${uid}" value="Sim" data-field="tIsolamento_${uid}"> Sim</label>
        <label><input type="radio" name="isolamento_${uid}" value="Não" data-field="tIsolamento_${uid}"> Não</label>
      </div>
      <div class="field" data-photo-key="tTubulacao_${uid}" data-photo-label="Tubulação / Isolamento"><label>Condição da tubulação e do isolamento (se aplicável)</label><textarea data-field="tCondicao_${uid}"></textarea></div>

      <div class="yn" style="margin-top:6px;">
        <span style="font-size:12.5px; font-weight:600; color:var(--steel); text-transform:uppercase;">Possui reparo?</span>
        <label><input type="radio" name="reparo_${uid}" value="Sim" data-field="tReparo_${uid}"> Sim</label>
        <label><input type="radio" name="reparo_${uid}" value="Não" data-field="tReparo_${uid}"> Não</label>
      </div>
      <div class="field" data-photo-key="tReparoFoto_${uid}" data-photo-label="Reparo"><label>Condição do reparo</label><textarea data-field="tCondicaoReparo_${uid}"></textarea></div>

      <div class="field" data-photo-key="tDescricao_${uid}" data-photo-label="Descrição da tubulação"><label>Descrição da tubulação</label><textarea data-field="tDescricaoTubulacao_${uid}" placeholder="Descreva o trecho, material, diâmetro, condições gerais observadas…"></textarea></div>

      <div class="field" data-photo-key="tValvulas_${uid}" data-photo-label="Válvulas e conexões"><label>Válvulas e conexões</label><textarea data-field="tValvulas_${uid}"></textarea></div>
      <div class="field" data-photo-key="tLigacoes_${uid}" data-photo-label="Ligações"><label>Ligações (flangeadas, roscadas e/ou soldadas)</label><textarea data-field="tLigacoes_${uid}"></textarea></div>
      <div class="field" data-photo-key="tSuportes_${uid}" data-photo-label="Suportes"><label>Suportes</label><textarea data-field="tSuportes_${uid}"></textarea></div>
      <div class="field" data-photo-key="tPassarelas_${uid}" data-photo-label="Passarelas"><label>Passarelas</label><textarea data-field="tPassarelas_${uid}"></textarea></div>
    </div>
  </div>`;
}

function addIsometrico(forceUid) {
  let uid;
  if (forceUid != null) { uid = forceUid; isoSeq = Math.max(isoSeq, uid); }
  else { isoSeq++; uid = isoSeq; }

  const wrap = document.createElement('div');
  wrap.innerHTML = isoCardHTML(uid).trim();
  const card = wrap.firstElementChild;
  document.getElementById('isoList').appendChild(card);
  buildItemPhotoWidgets(card);
  card.querySelector('[data-iso-remove]').addEventListener('click', () => {
    if (document.querySelectorAll('.iso-card').length <= 1) {
      alert('Mantenha pelo menos um isométrico.');
      return;
    }
    card.querySelectorAll('[data-photo-key]').forEach(h => delete itemPhotos[h.dataset.photoKey]);
    card.remove();
    relabelIsometricos();
    renderLinkedGallery();
    updateStatus();
  });
  relabelIsometricos();
  updateStatus();
  return card;
}

function relabelIsometricos() {
  const isoList = document.getElementById('isoList');
  const sectionCard = isoList.closest('section.card');
  const secNum = sectionCard ? sectionCard.querySelector('.num').textContent : '';
  document.querySelectorAll('.iso-card').forEach((card, i) => {
    card.querySelector('[data-iso-badge]').textContent = secNum ? `${secNum}.${i + 1}` : `${i + 1}`;
  });
}

document.getElementById('addIsoBtn').addEventListener('click', () => addIsometrico());

// ---------------------------------------------------------------------
// Fotos vinculadas a itens da inspeção
// ---------------------------------------------------------------------

function buildItemPhotoWidgets(root) {
  root = root || document;
  root.querySelectorAll('[data-photo-key]').forEach(host => {
    if (host.querySelector(':scope > .item-photos')) return; // já construído
    const key = host.dataset.photoKey;
    const label = host.dataset.photoLabel || key;
    itemPhotos[key] = itemPhotos[key] || [];

    const wrap = document.createElement('div');
    wrap.className = 'item-photos';
    const inputId = 'itemphoto-' + key;
    wrap.innerHTML = `
      <label class="item-photo-btn" for="${inputId}">📷 Foto — ${label} <span class="cnt" data-item-cnt="${key}"></span></label>
      <input class="item-photo-input" id="${inputId}" type="file" accept="image/*" capture="environment" multiple>
      <div class="item-thumbs" data-item-thumbs="${key}"></div>
    `;
    host.appendChild(wrap);

    wrap.querySelector('input[type=file]').addEventListener('change', (evt) => {
      const files = Array.from(evt.target.files || []);
      files.forEach(file => {
        const reader = new FileReader();
        reader.onload = () => {
          itemPhotos[key].push({ id: 'ip' + Date.now() + Math.random().toString(36).slice(2, 6), src: reader.result });
          renderItemThumbs(key);
          renderLinkedGallery();
          updateStatus();
        };
        reader.readAsDataURL(file);
      });
      evt.target.value = '';
    });
  });
}

function renderItemThumbs(key) {
  document.querySelectorAll(`[data-item-thumbs="${key}"]`).forEach(container => {
    container.innerHTML = '';
    (itemPhotos[key] || []).forEach(p => {
      const t = document.createElement('div');
      t.className = 'item-thumb';
      t.innerHTML = `<img src="${p.src}" alt=""><button data-rm-key="${key}" data-rm-id="${p.id}">✕</button>`;
      container.appendChild(t);
    });
  });
  document.querySelectorAll(`[data-item-cnt="${key}"]`).forEach(c => {
    const n = (itemPhotos[key] || []).length;
    c.textContent = n; c.classList.toggle('show', n > 0);
  });
  document.querySelectorAll(`[data-rm-key="${key}"]`).forEach(btn => {
    btn.onclick = () => {
      itemPhotos[key] = itemPhotos[key].filter(p => p.id !== btn.dataset.rmId);
      renderItemThumbs(key); renderLinkedGallery(); updateStatus();
    };
  });
}

function renderAllItemThumbs() {
  Object.keys(itemPhotos).forEach(renderItemThumbs);
}

function renderLinkedGallery() {
  const gallery = document.getElementById('linkedGallery');
  const groups = Object.entries(itemPhotos).filter(([k, v]) => v.length > 0);
  if (groups.length === 0) {
    gallery.innerHTML = '<div class="hint" style="color:var(--muted); font-size:12px;">Nenhuma foto anexada a itens ainda — use o botão "＋ Foto" em cada item da inspeção acima.</div>';
    return;
  }
  gallery.innerHTML = '';
  groups.forEach(([key, list]) => {
    const host = document.querySelector(`[data-photo-key="${key}"]`);
    const label = host ? (host.dataset.photoLabel || key) : key;
    const g = document.createElement('div');
    g.className = 'linked-group';
    g.innerHTML = `<div class="lg-title">${label} · ${list.length} foto${list.length > 1 ? 's' : ''}</div>
      <div class="item-thumbs">${list.map(p => `<div class="item-thumb"><img src="${p.src}" alt=""></div>`).join('')}</div>`;
    gallery.appendChild(g);
  });
}

// ---------------------------------------------------------------------
// Limpar / Novo
// ---------------------------------------------------------------------

function resetFormDom() {
  document.querySelectorAll('input[data-field]:not([readonly]), textarea[data-field]').forEach(el => {
    if (el.type === 'radio') el.checked = false; else el.value = '';
  });
  photos = [];
  Object.keys(itemPhotos).forEach(k => itemPhotos[k] = []);
  document.getElementById('isoList').innerHTML = '';
}

function clearAllBtnHandler() {
  if (!confirm('Limpar todos os campos e fotos preenchidos?')) return;
  startNewInspection();
}

function startNewInspection() {
  resetFormDom();
  currentRecordId = null;
  isoSeq = 0;
  addIsometrico();
  renderPhotos(); renderAllItemThumbs(); renderLinkedGallery(); syncPlate();
  setType('interna');
  updateStatus();
}

document.getElementById('clearAllBtn').addEventListener('click', clearAllBtnHandler);

// ---------------------------------------------------------------------
// Coleta de dados do formulário para o relatório .docx
// ---------------------------------------------------------------------

function inActiveTypeScope(el) {
  const wrap = el.closest('[data-types-only]');
  if (!wrap) return true;
  return wrap.dataset.typesOnly.split(',').includes(currentType);
}

function collectFieldsInto(root, block, skipEls) {
  skipEls = skipEls || [];
  const seen = new Set(skipEls);

  root.querySelectorAll('.yn').forEach(yn => {
    if (!inActiveTypeScope(yn)) return;
    const labelSpan = yn.querySelector('span');
    const label = labelSpan ? labelSpan.textContent.trim() : '';
    const checked = yn.querySelector('input[type=radio]:checked');
    block.rows.push({ label, value: checked ? checked.value : '' });
    yn.querySelectorAll('input[data-field]').forEach(i => seen.add(i));
  });

  root.querySelectorAll('.pressure-row').forEach(pr => {
    if (!inActiveTypeScope(pr)) return;
    const inputs = pr.querySelectorAll('input[data-field]');
    const unit = pr.querySelector('.unit-tag');
    const l1 = pr.querySelector('.field label');
    const label = (l1 ? l1.textContent.trim() : '') + (unit ? ' (' + unit.textContent.trim() + ')' : '');
    const kpa = inputs[0] ? inputs[0].value : '';
    const kgf = inputs[1] ? inputs[1].value : '';
    const value = [kpa && (kpa + ' kPa'), kgf && (kgf + ' kgf/cm²')].filter(Boolean).join(' / ');
    block.rows.push({ label, value });
    inputs.forEach(i => seen.add(i));
  });

  root.querySelectorAll('input[data-field], textarea[data-field], select[data-field]').forEach(el => {
    if (seen.has(el)) return;
    if (!inActiveTypeScope(el)) return;
    let label = '';
    const fieldWrap = el.closest('.field');
    if (fieldWrap) {
      const lab = fieldWrap.querySelector('label');
      if (lab) label = lab.textContent.trim();
    }
    if (!label && el.tagName === 'INPUT' && el.placeholder) label = el.placeholder;
    if (!label) label = el.dataset.field;
    block.rows.push({ label, value: el.value || '' });
    seen.add(el);
  });
}

function getPhotoKeyLabel(key) {
  const host = document.querySelector(`[data-photo-key="${key}"]`);
  if (!host) return key;
  let label = host.dataset.photoLabel || key;
  const isoCard = host.closest('.iso-card');
  if (isoCard) {
    const badge = isoCard.querySelector('[data-iso-badge]');
    const idInput = isoCard.querySelector('.iso-head input[data-field]');
    const isoTxt = 'Isométrico ' + (badge ? badge.textContent : '') + (idInput && idInput.value ? ' — ' + idInput.value : '');
    label = isoTxt + ' · ' + label;
  }
  return label;
}

function collectReportData() {
  const sections = [];
  document.querySelectorAll('[data-section]:not(.hidden-type)').forEach(sec => {
    const title = sec.querySelector('h2').textContent.replace(/^\d+(\.\d+)?/, '').trim();
    const secObj = { title, blocks: [] };
    const isoList = sec.querySelector('#isoList');
    const pg = sec.querySelector('#photoGrid');

    if (isoList && currentType === 'tubulacao') {
      sec.querySelectorAll('.iso-card').forEach(card => {
        const badge = card.querySelector('[data-iso-badge]').textContent;
        const idInput = card.querySelector('.iso-head input[data-field]');
        const idVal = idInput ? idInput.value : '';
        const block = { heading: 'Isométrico ' + badge + (idVal ? ' — ' + idVal : ''), rows: [], photos: [] };
        collectFieldsInto(card, block, idInput ? [idInput] : []);
        secObj.blocks.push(block);
      });
    } else if (pg) {
      // Seção "Registro Fotográfico": consolida fotos vinculadas a itens + fotos gerais
      const block = { rows: [], photos: [] };
      collectFieldsInto(sec, block, []);
      secObj.blocks.push(block);

      Object.keys(itemPhotos).forEach(key => {
        const list = itemPhotos[key] || [];
        if (!list.length) return;
        const host = document.querySelector(`[data-photo-key="${key}"]`);
        if (host && (host.closest('.hidden-type') || !inActiveTypeScope(host))) return; // ignora itens de outros tipos de inspeção
        const label = getPhotoKeyLabel(key);
        secObj.blocks.push({
          heading: label,
          rows: [],
          photos: list.map(p => ({ src: p.src, caption: label }))
        });
      });

      if (photos.length) {
        secObj.blocks.push({
          heading: 'Fotos gerais',
          rows: [],
          photos: photos.map(p => ({ src: p.src, caption: p.caption || 'Foto geral' }))
        });
      }
    } else {
      const block = { rows: [], photos: [] };
      collectFieldsInto(sec, block, []);
      secObj.blocks.push(block);
    }
    sections.push(secObj);
  });
  return sections;
}

// ---------------------------------------------------------------------
// Exportar como Word (.docx)
// ---------------------------------------------------------------------

// Em aparelhos corporativos com armazenamento interno bloqueado (política de
// proteção contra vazamento de dados), a saída de arquivos costuma ser feita
// só pelos apps oficiais (e-mail/Teams corporativo). Por isso, sempre que o
// aparelho suportar, abrimos direto o menu "Compartilhar" nativo — ele deixa
// escolher o app corporativo e anexa o .docx sem passar pelo armazenamento
// interno. Só cai no download tradicional se o compartilhamento não existir.
async function shareOrDownloadDocx(blob, filename) {
  try {
    if (navigator.canShare) {
      const file = new File([blob], filename, { type: blob.type });
      if (navigator.canShare({ files: [file] })) {
        await navigator.share({ files: [file], title: filename });
        return;
      }
    }
  } catch (err) {
    if (err && err.name === 'AbortError') return; // usuário cancelou o compartilhamento
    console.warn('Compartilhamento indisponível, baixando o arquivo em vez disso:', err);
  }

  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

async function exportWord() {
  const btn = document.getElementById('exportBtn');
  const originalLabel = btn.textContent;
  btn.textContent = 'Gerando Word…'; btn.disabled = true;
  try {
    const sections = collectReportData();
    const t = TYPES[currentType];
    const tag = document.querySelector('[data-field="tag"]').value || '';
    const unidade = document.querySelector('[data-field="unidade"]').value || '';
    const local = document.querySelector('[data-field="local"]').value || '';

    const blob = await buildDocxBlob({ typeLabel: t.label, tag, unidade, local, sections });

    const tagSafe = (tag || 'relatorio').replace(/[^a-zA-Z0-9-_]/g, '_');
    const dateForFile = new Date().toISOString().slice(0, 10);
    const filename = 'relatorio_' + currentType + '_' + tagSafe + '_' + dateForFile + '.docx';

    await shareOrDownloadDocx(blob, filename);
  } catch (err) {
    console.error(err);
    alert('Não foi possível gerar o arquivo Word: ' + err.message);
  } finally {
    btn.textContent = originalLabel; btn.disabled = false;
  }
}

document.getElementById('exportBtn').addEventListener('click', exportWord);
document.getElementById('printBtn').addEventListener('click', () => window.print());

// ---------------------------------------------------------------------
// Serializar / restaurar formulário (para salvar e reabrir no Histórico)
// ---------------------------------------------------------------------

function serializeForm() {
  const fields = {};
  document.querySelectorAll('[data-field]').forEach(el => {
    const key = el.dataset.field;
    if (key === 'tipoInspecaoTxt') return; // derivado do tipo, não precisa persistir
    if (el.type === 'radio') {
      if (!(key in fields)) fields[key] = '';
      if (el.checked) fields[key] = el.value;
    } else {
      fields[key] = el.value;
    }
  });
  const isometricos = Array.from(document.querySelectorAll('.iso-card')).map(card => ({ uid: Number(card.dataset.iso) }));
  return {
    fields,
    isometricos,
    photosGeral: photos.map(p => ({ ...p })),
    itemPhotos: JSON.parse(JSON.stringify(itemPhotos))
  };
}

function populateForm(record) {
  resetFormDom();
  isoSeq = 0;

  const uids = (record.isometricos || []).map(x => x.uid);
  if (record.tipo === 'tubulacao' && uids.length) {
    uids.forEach(uid => addIsometrico(uid));
  } else {
    addIsometrico();
  }

  setType(record.tipo);

  Object.entries(record.fields || {}).forEach(([key, val]) => {
    if (key === 'tipoInspecaoTxt') return;
    document.querySelectorAll(`[data-field="${CSS.escape(key)}"]`).forEach(el => {
      if (el.type === 'radio') { el.checked = (el.value === val && val !== ''); }
      else { el.value = val; }
    });
  });

  photos = (record.photosGeral || []).map(p => ({ ...p }));
  Object.keys(itemPhotos).forEach(k => itemPhotos[k] = []);
  Object.entries(record.itemPhotos || {}).forEach(([k, v]) => {
    itemPhotos[k] = (v || []).map(p => ({ ...p }));
  });

  renderPhotos();
  renderAllItemThumbs();
  renderLinkedGallery();
  syncPlate();
  updateStatus();
}

// ---------------------------------------------------------------------
// Salvar inspeção (IndexedDB)
// ---------------------------------------------------------------------

function uuid() {
  if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
  return 'id-' + Date.now() + '-' + Math.random().toString(36).slice(2, 10);
}

async function saveInspection() {
  const tagVal = (document.querySelector('[data-field="tag"]').value || '').trim();
  if (!tagVal) {
    alert('Preencha ao menos a TAG do equipamento antes de salvar.');
    return;
  }
  const btn = document.getElementById('saveBtn');
  const originalLabel = btn.textContent;
  btn.disabled = true; btn.textContent = 'Salvando…';
  try {
    const now = new Date().toISOString();
    const serialized = serializeForm();
    let createdAt = now;
    if (currentRecordId) {
      const existing = await dbGetInspection(currentRecordId);
      if (existing && existing.createdAt) createdAt = existing.createdAt;
    }
    const record = {
      id: currentRecordId || uuid(),
      tipo: currentType,
      tag: serialized.fields.tag || '',
      unidade: serialized.fields.unidade || '',
      local: serialized.fields.local || '',
      numRI: serialized.fields.numRI || '',
      dataInicio: serialized.fields.dataInicio || '',
      dataFim: serialized.fields.dataFim || '',
      createdAt,
      savedAt: now,
      fields: serialized.fields,
      isometricos: serialized.isometricos,
      photosGeral: serialized.photosGeral,
      itemPhotos: serialized.itemPhotos
    };
    await dbSaveInspection(record);
    currentRecordId = record.id;
    btn.textContent = 'Salvo ✓';
    setTimeout(() => { btn.textContent = originalLabel; }, 1600);
  } catch (err) {
    console.error(err);
    alert('Não foi possível salvar a inspeção: ' + err.message);
    btn.textContent = originalLabel;
  } finally {
    btn.disabled = false;
  }
}

document.getElementById('saveBtn').addEventListener('click', saveInspection);

// ---------------------------------------------------------------------
// Histórico de inspeções
// ---------------------------------------------------------------------

function formatDateBR(isoDate) {
  if (!isoDate) return '—';
  const parts = String(isoDate).split('-');
  if (parts.length === 3) return `${parts[2]}/${parts[1]}/${parts[0]}`;
  return isoDate;
}

function formatDateTimeBR(isoDateTime) {
  if (!isoDateTime) return '—';
  const d = new Date(isoDateTime);
  if (isNaN(d.getTime())) return '—';
  return d.toLocaleDateString('pt-BR') + ' ' + d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
}

async function renderHistoryView() {
  const container = document.getElementById('historyList');
  container.innerHTML = '<div class="history-empty">Carregando…</div>';
  let records;
  try {
    records = await dbGetAllInspections();
  } catch (err) {
    container.innerHTML = '<div class="history-empty">Não foi possível ler o histórico salvo neste aparelho.</div>';
    console.error(err);
    return;
  }
  records.sort((a, b) => (b.savedAt || '').localeCompare(a.savedAt || ''));

  if (!records.length) {
    container.innerHTML = '<div class="history-empty">Nenhuma inspeção salva ainda. Preencha um relatório no Formulário e clique em "Salvar inspeção".</div>';
    return;
  }

  container.innerHTML = records.map(r => {
    const t = TYPES[r.tipo] || { label: r.tipo, color: '#666' };
    return `<div class="history-row" data-id="${r.id}">
      <div class="hr-info">
        <div class="hr-tag">${escapeHtml(r.tag || '(sem TAG)')}</div>
        <span class="hr-badge" style="background:${t.color}">${escapeHtml(t.label)}</span>
        <div class="hr-meta">Data da inspeção: <b>${formatDateBR(r.dataInicio)}</b> · Unidade/Local: ${escapeHtml(r.unidade || '—')} / ${escapeHtml(r.local || '—')} · Salvo em ${formatDateTimeBR(r.savedAt)}</div>
      </div>
      <div class="hr-actions">
        <button class="btn ghost" data-action="open" data-id="${r.id}">Reabrir</button>
        <button class="btn primary" data-action="docx" data-id="${r.id}">Gerar Word</button>
        <button class="btn ghost" data-action="delete" data-id="${r.id}" style="color:var(--red);">Excluir</button>
      </div>
    </div>`;
  }).join('');

  container.querySelectorAll('[data-action="open"]').forEach(btn => {
    btn.addEventListener('click', () => openInspectionFromHistory(btn.dataset.id));
  });
  container.querySelectorAll('[data-action="docx"]').forEach(btn => {
    btn.addEventListener('click', () => regenerateDocxFromHistory(btn.dataset.id, btn));
  });
  container.querySelectorAll('[data-action="delete"]').forEach(btn => {
    btn.addEventListener('click', () => deleteInspectionFromHistory(btn.dataset.id));
  });
}

function escapeHtml(str) {
  return String(str == null ? '' : str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

async function openInspectionFromHistory(id) {
  const record = await dbGetInspection(id);
  if (!record) { alert('Inspeção não encontrada.'); return; }
  currentRecordId = record.id;
  populateForm(record);
  showFormView();
}

async function regenerateDocxFromHistory(id, btn) {
  const record = await dbGetInspection(id);
  if (!record) { alert('Inspeção não encontrada.'); return; }
  const originalLabel = btn.textContent;
  btn.disabled = true; btn.textContent = 'Gerando…';
  try {
    currentRecordId = record.id;
    populateForm(record);
    await exportWord();
    showFormView();
  } finally {
    btn.disabled = false; btn.textContent = originalLabel;
  }
}

async function deleteInspectionFromHistory(id) {
  if (!confirm('Excluir esta inspeção salva? Essa ação não pode ser desfeita.')) return;
  await dbDeleteInspection(id);
  renderHistoryView();
}

// ---------------------------------------------------------------------
// Service worker (PWA offline)
// ---------------------------------------------------------------------

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('sw.js').catch(err => console.error('Falha ao registrar service worker:', err));
  });
}

// ---------------------------------------------------------------------
// Inicialização
// ---------------------------------------------------------------------

buildItemPhotoWidgets(document);
addIsometrico();
setType('interna');
showFormView();
