import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
const root = fileURLToPath(new URL('../', import.meta.url));
const read = file => readFile(path.join(root, file), 'utf8');
export async function build() {
  const [template, styles, core, client, server, manifest] = await Promise.all(['src/index.html','src/styles.css','src/core.js','src/app.js','src/server.gs','src/appsscript.json'].map(read));
  const html = template.replace('/* APP_STYLES */', styles).replace('/* APP_CORE */', core.replaceAll('</script', '<\\/script')).replace('/* APP_CLIENT */', client.replaceAll('</script','<\\/script'));
  await mkdir(path.join(root,'dist'), {recursive:true});
  await Promise.all([
    writeFile(path.join(root,'dist','Index.html'), html),
    writeFile(path.join(root,'dist','Code.gs'), core+'\n'+server),
    writeFile(path.join(root,'dist','appsscript.json'), JSON.stringify(JSON.parse(manifest),null,2)+'\n'),
    writeFile(path.join(root,'dist','index.html'), html)
  ]);
  return html;
}
if (process.argv[1] === fileURLToPath(import.meta.url)) { const html = await build(); console.log(`Apps Script用のファイルをdistへ作成しました（画面 ${Math.round(Buffer.byteLength(html)/1024)} KB）。`); }
