// Tela Arquivo (Backstage), como no Word.
import { h, clear } from './dom.js';
import { icon } from './icons.js';

const FORMATS = [
  { ext: 'docx', label: 'Documento do Word (*.docx)', icon: 'word16' },
  { ext: 'odt', label: 'Texto OpenDocument (*.odt)', icon: 'newdoc16' },
  { ext: 'md', label: 'Markdown (*.md)', icon: 'newdoc16' },
  { ext: 'html', label: 'Página da Web (*.html)', icon: 'webLayout16' },
  { ext: 'txt', label: 'Texto sem Formatação (*.txt)', icon: 'newdoc16' },
  { ext: 'rtf', label: 'Rich Text Format (*.rtf)', icon: 'newdoc16' },
  { ext: 'pdf', label: 'PDF (*.pdf)', icon: 'pdf16' },
];

function when(ms) {
  const d = new Date(ms);
  const now = new Date();
  const days = Math.floor((now - d) / 86400000);
  if (d.toDateString() === now.toDateString()) return 'Hoje, ' + d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
  if (days < 2) return 'Ontem';
  if (days < 7) return 'Há ' + days + ' dias';
  return d.toLocaleDateString('pt-BR');
}

function fileIcon(name) {
  const ext = (name.split('.').pop() || '').toLowerCase();
  if (ext === 'docx' || ext === 'doc' || ext === 'docm' || ext === 'dotx') return icon('word16');
  if (ext === 'pdf') return icon('pdf16');
  return icon('newdoc16');
}

export class Backstage {
  constructor(app) {
    this.app = app;
    this.section = 'home';
    this.nav = h('nav', { class: 'bs-nav' });
    this.content = h('div', { class: 'bs-content' });
    this.el = h('div', { class: 'backstage', hidden: true }, this.nav, this.content);
    const item = (id, label, ic) => {
      const b = h('button', { class: 'bs-item', type: 'button', dataset: { id } }, ic ? h('span', { class: 'bs-ic', html: icon(ic) }) : null, h('span', { text: label }));
      b.addEventListener('click', () => this.show(id));
      return b;
    };
    this.nav.append(
      h('button', { class: 'bs-back', type: 'button', title: 'Voltar', html: icon('back'), onclick: () => this.close() }),
      item('home', 'Início', 'home16'), item('new', 'Novo', 'newdoc16'), item('open', 'Abrir', 'folder16'),
      h('div', { class: 'bs-sep' }),
      item('info', 'Informações'), item('save', 'Salvar'), item('saveas', 'Salvar como'), item('print', 'Imprimir'), item('export', 'Exportar'), item('close', 'Fechar'),
      h('div', { class: 'bs-spacer' }),
      item('options', 'Opções'));
    this.onKey = (e) => { if (e.key === 'Escape') { e.preventDefault(); this.close(); } };
  }

  open(section = 'home') {
    this.el.hidden = false;
    document.addEventListener('keydown', this.onKey);
    this.app.setOverlayColor(true);
    this.show(section);
  }

  close() {
    this.el.hidden = true;
    document.removeEventListener('keydown', this.onKey);
    this.app.setOverlayColor(false);
    this.app.focusEditor();
  }

  get isOpen() { return !this.el.hidden; }

  async show(id) {
    const app = this.app;
    if (id === 'save') { this.close(); app.save(); return; }
    if (id === 'close') { this.close(); app.closeWindow(); return; }
    if (id === 'options') { this.close(); app.dialog('options'); return; }
    this.section = id;
    for (const b of this.nav.querySelectorAll('.bs-item')) b.classList.toggle('active', b.dataset.id === id);
    clear(this.content);
    const title = (t) => h('h1', { class: 'bs-title', text: t });
    if (id === 'home' || id === 'new') {
      const hour = new Date().getHours();
      this.content.appendChild(title(id === 'home' ? (hour >= 5 && hour < 12 ? 'Bom dia' : hour >= 12 && hour < 18 ? 'Boa tarde' : 'Boa noite') : 'Novo'));
      const tiles = h('div', { class: 'bs-tiles' },
        h('button', { class: 'bs-tile', type: 'button', onclick: () => { this.close(); app.newDocument(); } }, h('div', { class: 'bs-thumb blank' }), h('span', { text: 'Documento em branco' })),
        h('button', { class: 'bs-tile', type: 'button', onclick: () => { this.close(); app.newDocument('abnt'); } }, h('div', { class: 'bs-thumb abnt' }, h('div', { class: 'bs-thumb-lines' })), h('span', { text: 'Trabalho acadêmico (ABNT)' })),
        h('button', { class: 'bs-tile', type: 'button', onclick: () => { this.close(); app.newDocument('md'); } }, h('div', { class: 'bs-thumb md' }, h('span', { text: 'M↓' })), h('span', { text: 'Documento Markdown' })));
      this.content.appendChild(tiles);
      if (id === 'home') {
        this.content.appendChild(h('h2', { class: 'bs-sub', text: 'Recentes' }));
        this.content.appendChild(await this.recentList());
      }
      return;
    }
    if (id === 'open') {
      this.content.appendChild(title('Abrir'));
      const browse = h('button', { class: 'bs-browse', type: 'button', onclick: () => { this.close(); app.openDialog(); } }, h('span', { class: 'bs-ic', html: icon('folder16') }), h('span', { text: 'Procurar' }));
      this.content.appendChild(h('div', { class: 'bs-open' }, h('div', { class: 'bs-open-left' }, browse), h('div', { class: 'bs-open-right' }, h('h2', { class: 'bs-sub', text: 'Documentos' }), await this.recentList())));
      return;
    }
    if (id === 'info') {
      this.content.appendChild(title('Informações'));
      const st = app.stats(true);
      const p = app.model.properties || {};
      const row = (k, v) => h('tr', {}, h('td', { class: 'k', text: k }), h('td', { text: v }));
      const tbl = h('table', { class: 'bs-props' },
        row('Nome', app.fileName()),
        row('Local', app.filePath || 'Não salvo'),
        row('Formato', app.formatLabel()),
        row('Páginas', String(st.pages)),
        row('Palavras', st.words.toLocaleString('pt-BR')),
        row('Caracteres', st.chars.toLocaleString('pt-BR')),
        row('Autor', p.author || '-'),
        row('Última modificação por', p.lastModifiedBy || '-'),
        row('Criado', p.created ? new Date(p.created).toLocaleString('pt-BR') : '-'),
        row('Modificado', p.modified ? new Date(p.modified).toLocaleString('pt-BR') : '-'));
      const actions = h('div', { class: 'bs-actions' });
      if (app.filePath) actions.appendChild(h('button', { class: 'btn', type: 'button', text: 'Abrir Local do Arquivo', onclick: () => app.host.showInFolder(app.filePath) }));
      this.content.append(actions, tbl);
      return;
    }
    if (id === 'saveas') {
      this.content.appendChild(title('Salvar como'));
      const list = h('div', { class: 'bs-formats' });
      for (const f of FORMATS) {
        list.appendChild(h('button', { class: 'bs-format', type: 'button', onclick: () => { this.close(); if (f.ext === 'pdf') app.exportPdf(); else app.saveAs(f.ext); } }, h('span', { class: 'bs-ic', html: icon(f.icon) }), h('span', { text: f.label })));
      }
      this.content.appendChild(list);
      return;
    }
    if (id === 'print') {
      this.content.appendChild(title('Imprimir'));
      const btn = h('button', { class: 'bs-print-btn', type: 'button', onclick: () => { this.close(); app.print(); } }, h('span', { html: icon('print') }), h('span', { text: 'Imprimir' }));
      const pdf = h('button', { class: 'bs-print-btn', type: 'button', onclick: () => { this.close(); app.exportPdf(); } }, h('span', { html: icon('pdf') }), h('span', { text: 'Salvar como PDF' }));
      const preview = h('div', { class: 'bs-preview' });
      this.content.append(h('div', { class: 'bs-print' }, h('div', { class: 'bs-print-left' }, btn, pdf, h('p', { class: 'muted', text: 'A impressora, as cópias e o intervalo de páginas são escolhidos na janela de impressão do Windows.' })), preview));
      app.renderPreview(preview);
      return;
    }
    if (id === 'export') {
      this.content.appendChild(title('Exportar'));
      const list = h('div', { class: 'bs-formats' },
        h('button', { class: 'bs-format big', type: 'button', onclick: () => { this.close(); app.exportPdf(); } }, h('span', { class: 'bs-ic', html: icon('pdf') }), h('span', {}, h('b', { text: 'Criar Documento PDF' }), h('br'), h('span', { class: 'muted', text: 'Preserva o layout, a formatação, as fontes e as imagens.' }))),
        h('h2', { class: 'bs-sub', text: 'Alterar Tipo de Arquivo' }));
      for (const f of FORMATS.filter((x) => x.ext !== 'pdf')) {
        list.appendChild(h('button', { class: 'bs-format', type: 'button', onclick: () => { this.close(); app.saveAs(f.ext); } }, h('span', { class: 'bs-ic', html: icon(f.icon) }), h('span', { text: f.label })));
      }
      this.content.appendChild(list);
    }
  }

  async recentList() {
    const app = this.app;
    const list = h('div', { class: 'bs-recent' });
    const recent = await app.host.recent();
    if (!recent.length) {
      list.appendChild(h('p', { class: 'muted', text: 'Você não abriu nenhum documento recentemente. Clique em Abrir para procurar um arquivo.' }));
      return list;
    }
    list.appendChild(h('div', { class: 'bs-recent-head' }, h('span', { text: 'Nome' }), h('span', { text: 'Data de modificação' })));
    for (const r of recent) {
      const row = h('div', { class: 'bs-recent-row', title: r.path },
        h('span', { class: 'bs-ic', html: fileIcon(r.name) }),
        h('span', { class: 'bs-rn' }, h('span', { class: 'bs-rname', text: r.name }), h('span', { class: 'bs-rpath', text: r.folder })),
        h('span', { class: 'bs-rdate', text: when(r.time) }),
        h('button', { class: 'bs-rx', type: 'button', title: 'Remover da lista', html: icon('close') }));
      row.addEventListener('click', (e) => {
        if (e.target.closest('.bs-rx')) {
          e.stopPropagation();
          app.host.removeRecent(r.path).then(() => this.show(this.section));
          return;
        }
        this.close();
        app.openPath(r.path);
      });
      list.appendChild(row);
    }
    return list;
  }
}
