// Backup opcional no Google Drive PESSOAL da usuária.
//
// Isso é totalmente à parte do funcionamento normal do app, que continua
// 100% offline e local (IndexedDB) sem depender disso em nenhum momento.
// Só entra em ação quando a pessoa clica em "Backup no Google Drive" ou
// "Restaurar do Google Drive" — e exige internet e login Google nesse
// instante. Usamos o escopo mínimo (drive.file), que só permite ao app ver
// e criar arquivos que ele mesmo criou — nunca o resto do Drive da pessoa.
//
// Para habilitar, preencha o Client ID OAuth abaixo (gratuito, criado em
// console.cloud.google.com). Sem isso, os botões avisam que a função ainda
// não foi configurada.
const GOOGLE_DRIVE_CLIENT_ID = ''; // <-- cole aqui algo como "123456789-abc.apps.googleusercontent.com"
const GOOGLE_DRIVE_SCOPE = 'https://www.googleapis.com/auth/drive.file';
const GOOGLE_DRIVE_FOLDER_NAME = 'Inspecoes App - Backups';

let __gisLoaded = false;
let __tokenClient = null;
let __driveAccessToken = null;

function ensureGoogleIdentity() {
  if (__gisLoaded && window.google && window.google.accounts) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = 'https://accounts.google.com/gsi/client';
    s.onload = () => { __gisLoaded = true; resolve(); };
    s.onerror = () => reject(new Error('Não foi possível carregar o login do Google (verifique a internet).'));
    document.head.appendChild(s);
  });
}

function getDriveAccessToken() {
  return new Promise((resolve, reject) => {
    if (!GOOGLE_DRIVE_CLIENT_ID) {
      reject(new Error('O backup no Google Drive ainda não foi configurado neste app.'));
      return;
    }
    ensureGoogleIdentity().then(() => {
      if (!__tokenClient) {
        __tokenClient = google.accounts.oauth2.initTokenClient({
          client_id: GOOGLE_DRIVE_CLIENT_ID,
          scope: GOOGLE_DRIVE_SCOPE,
          callback: () => {} // definido a cada chamada, logo abaixo
        });
      }
      __tokenClient.callback = (resp) => {
        if (resp.error) { reject(new Error('Login do Google cancelado ou negado.')); return; }
        __driveAccessToken = resp.access_token;
        resolve(__driveAccessToken);
      };
      __tokenClient.requestAccessToken({ prompt: __driveAccessToken ? '' : 'consent' });
    }).catch(reject);
  });
}

async function driveApiFetch(url, options) {
  options = options || {};
  options.headers = Object.assign({ Authorization: 'Bearer ' + __driveAccessToken }, options.headers || {});
  const resp = await fetch(url, options);
  if (!resp.ok) {
    const text = await resp.text().catch(() => '');
    throw new Error('Erro na API do Google Drive (' + resp.status + '): ' + text.slice(0, 200));
  }
  return resp;
}

async function findOrCreateDriveFolder() {
  const q = encodeURIComponent(`name='${GOOGLE_DRIVE_FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and trashed=false`);
  const searchResp = await driveApiFetch(`https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)`);
  const searchData = await searchResp.json();
  if (searchData.files && searchData.files.length) return searchData.files[0].id;

  const createResp = await driveApiFetch('https://www.googleapis.com/drive/v3/files', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name: GOOGLE_DRIVE_FOLDER_NAME, mimeType: 'application/vnd.google-apps.folder' })
  });
  const created = await createResp.json();
  return created.id;
}

async function uploadJsonToDriveFolder(folderId, filename, jsonString) {
  const metadata = { name: filename, parents: [folderId], mimeType: 'application/json' };
  const boundary = '-------inspecoesapp' + Date.now();
  const body =
    `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n` +
    `--${boundary}\r\nContent-Type: application/json\r\n\r\n${jsonString}\r\n` +
    `--${boundary}--`;

  await driveApiFetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
    method: 'POST',
    headers: { 'Content-Type': `multipart/related; boundary=${boundary}` },
    body
  });
}

async function listDriveBackupFiles(folderId) {
  const q = encodeURIComponent(`'${folderId}' in parents and trashed=false`);
  const resp = await driveApiFetch(`https://www.googleapis.com/drive/v3/files?q=${q}&orderBy=createdTime desc&fields=files(id,name,createdTime)`);
  const data = await resp.json();
  return data.files || [];
}

async function downloadDriveFileText(fileId) {
  const resp = await driveApiFetch(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`);
  return resp.text();
}
