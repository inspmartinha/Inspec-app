// Corrige a orientação das fotos (comum em fotos tiradas pela câmera do
// celular, que gravam um metadado EXIF de rotação junto do arquivo) antes de
// guardá-las. O navegador já usa esse metadado pra exibir a foto em pé na
// tela — mas o Word ignora esse metadado e mostra a foto "crua", de lado ou
// de cabeça para baixo. A correção é reaproveitar a própria decodificação já
// corrigida do navegador: desenhamos a imagem (já em pé) num canvas e
// reexportamos a partir dali — o arquivo gerado não carrega mais metadado de
// rotação nenhum, então fica certo tanto na prévia do app quanto no .docx.
//
// (Importante: não dá pra fazer essa rotação "na mão" a partir do valor EXIF
// bruto — navegadores atuais já aplicam a rotação sozinhos ao decodificar a
// imagem via <img>/canvas, então rotacionar de novo bagunçaria a foto.)

function readImageNormalized(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      try {
        const canvas = document.createElement('canvas');
        canvas.width = img.naturalWidth;
        canvas.height = img.naturalHeight;
        canvas.getContext('2d').drawImage(img, 0, 0);
        const isPng = file.type === 'image/png';
        resolve(isPng ? canvas.toDataURL('image/png') : canvas.toDataURL('image/jpeg', 0.92));
      } catch (err) {
        reject(err);
      } finally {
        URL.revokeObjectURL(url);
      }
    };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('Não foi possível carregar a foto.')); };
    img.src = url;
  });
}

function readFileAsDataUrlPlain(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

// Lê um arquivo de foto e devolve uma dataURL já "achatada" na orientação
// correta (sem depender de metadado). Só reprocessa JPEG (único formato de
// câmera com esse problema — PNG não tem esse metadado, então passa direto).
// Em qualquer erro no meio do caminho, cai de volta pro comportamento
// simples (lê o arquivo como está) em vez de travar a captura da foto.
async function readPhotoFile(file) {
  const isJpeg = file.type === 'image/jpeg' || file.type === 'image/jpg';
  if (!isJpeg) return readFileAsDataUrlPlain(file);

  try {
    return await readImageNormalized(file);
  } catch (err) {
    console.warn('Não foi possível normalizar a orientação da foto, usando original:', err);
    return readFileAsDataUrlPlain(file);
  }
}
