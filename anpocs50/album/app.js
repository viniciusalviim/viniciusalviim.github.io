/* ============================================================
   ÁLBUM DE FIGURINHAS ANPOCS 50 ANOS
   Conteúdo: planilha Modelo_Conteudo_Album_ANPOCS.xlsx
   Imagens: pasta cartazes/
   O progresso fica salvo no próprio aparelho (localStorage).
   ============================================================ */

const CONFIG = {
  planilha: "Modelo_Conteudo_Album_ANPOCS.xlsx",
  pastaCartazes: "cartazes/",
  figurinhasPorPacote: 5,
  porPagina: 5,
  chaveSalvamento: "album-anpocs-50"
};

let cartazes = [];
let perguntas = [];

// progresso salvo
let progresso = { obtidas: [], coladas: [], respostas: {} };

// estado da tela (não é salvo)
const tela = { pagina: 0, selecionada: null, recem: new Set() };

const app = document.getElementById("app");
const janela = document.getElementById("janela");
const toastEl = document.getElementById("toast");

/* ---------- utilidades ---------- */

const esc = (t) => String(t ?? "").replace(/[&<>"']/g, (c) => ({
  "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
}[c]));
const txt = (v) => String(v ?? "").trim();

let toastTimer = null;
function toast(msg) {
  toastEl.textContent = msg;
  toastEl.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { toastEl.hidden = true; }, 3200);
}

const cartazPorId = (id) => cartazes.find((c) => c.id === id);
const indiceCartaz = (id) => cartazes.findIndex((c) => c.id === id);
const paginasPorVez = () => (window.innerWidth >= 900 ? 2 : 1);
// páginas do álbum: primeiro os cartazes, depois as figurinhas especiais (cromadas) em páginas próprias
let paginasAlbum = [];
function montarPaginas() {
  const blocos = (lista) => {
    const r = [];
    for (let i = 0; i < lista.length; i += CONFIG.porPagina) r.push(lista.slice(i, i + CONFIG.porPagina));
    return r;
  };
  // cartazes em ordem e, no fim, as figurinhas especiais (cromadas)
  cartazes = [...cartazes.filter((c) => !c.cromada), ...cartazes.filter((c) => c.cromada)];
  cartazes.forEach((c, i) => { c.numero = i + 1; });
  paginasAlbum = blocos(cartazes);
}
const totalPaginas = () => paginasAlbum.length;

/* ---------- salvar / carregar progresso ---------- */

function salvar() {
  try { localStorage.setItem(CONFIG.chaveSalvamento, JSON.stringify(progresso)); } catch (e) {}
}

function lerProgresso() {
  try {
    const p = JSON.parse(localStorage.getItem(CONFIG.chaveSalvamento) || "null");
    if (!p) return;
    const ids = new Set(cartazes.map((c) => c.id));
    progresso.coladas = (p.coladas || []).filter((id) => ids.has(id));
    progresso.obtidas = (p.obtidas || []).filter((id) => ids.has(id) && !progresso.coladas.includes(id));
    progresso.respostas = p.respostas || {};
  } catch (e) {}
}

const temProgresso = () =>
  progresso.coladas.length + progresso.obtidas.length + Object.keys(progresso.respostas).length > 0;

/* ---------- planilha ---------- */

async function carregarPlanilha() {
  const r = await fetch(CONFIG.planilha + "?v=" + Date.now());
  if (!r.ok) throw new Error("Planilha não encontrada");
  const livro = XLSX.read(await r.arrayBuffer(), { type: "array" });
  const aba = (n) => (livro.Sheets[n] ? XLSX.utils.sheet_to_json(livro.Sheets[n], { defval: "" }) : []);

  cartazes = aba("Cartazes")
    .filter((l) => txt(l.id_cartaz) && (txt(l.ano) || txt(l.edicao)))
    .sort((a, b) => (Number(a.ordem_album) || 9999) - (Number(b.ordem_album) || 9999))
    .map((l, i) => {
      const ano = txt(l.ano);
      return {
        id: txt(l.id_cartaz),
        ano,
        rotulo: ano || txt(l.edicao),          // figurinhas brinde: deixe o ano vazio e use a edição
        edicao: txt(l.edicao),
        cidade: txt(l.cidade_local),
        tema: txt(l.tema_titulo),
        detalhes: txt(l.texto_detalhes),
        curiosidade: txt(l.curiosidade),
        fonteInfo: txt(l.fonte_informacoes),
        fonteImagem: txt(l.fonte_imagem),
        imagem: txt(l.arquivo_imagem) ? CONFIG.pastaCartazes + txt(l.arquivo_imagem) : "",
        hue: Number(ano) ? (Number(ano) * 37 + i) % 360 : (i * 47) % 360,
        cromada: ["sim", "s"].includes(txt(l.cromada).toLowerCase())
      };
    });

  perguntas = aba("Perguntas")
    .map((l) => ({
      id: txt(l.id_pergunta),
      enunciado: txt(l.pergunta),
      a: txt(l.alternativa_a), b: txt(l.alternativa_b),
      c: txt(l.alternativa_c), d: txt(l.alternativa_d),
      correta: txt(l.resposta_correta).toLowerCase(),
      explicacao: txt(l.explicacao),
      ativa: txt(l.ativa).toLowerCase()
    }))
    .filter((p) => p.ativa === "sim" && p.id && p.enunciado && p.a && p.b && p.c && p.d &&
      ["a", "b", "c", "d"].includes(p.correta));

  if (!cartazes.length) throw new Error("Nenhum cartaz na aba Cartazes");
  montarPaginas();
}

/* ---------- figurinha ---------- */

const numeroFig = (n) => String(n).padStart(2, "0");

function figurinhaHtml(c) {
  // cartaz que não preenche a moldura: o espaço sobrando é coberto por uma versão desfocada do próprio cartaz
  const fundo = c.imagem && !c.cromada
    ? `<div class="figurinha-fundo" style="--img:url('${esc(c.imagem)}')"></div>` : "";
  const arte = c.imagem
    ? `${fundo}<img src="${esc(c.imagem)}" alt="Cartaz ${esc(c.rotulo)}" data-provisoria>`
    : provisoriaHtml();
  return `<div class="figurinha ${c.cromada ? "cromada" : ""}" style="--hue:${c.hue}">${arte}<div class="figurinha-numero">${numeroFig(c.numero)}</div><div class="figurinha-ano">${esc(c.rotulo)}</div></div>`;
}
const provisoriaHtml = () => `<div class="figurinha-provisoria">ANPOCS<small>Cartaz provisório</small></div>`;

// imagem que não existe vira figurinha provisória
document.addEventListener("error", (e) => {
  const img = e.target;
  if (img.tagName === "IMG" && img.hasAttribute("data-provisoria")) {
    img.outerHTML = provisoriaHtml();
  }
}, true);

/* ---------- tela inicial ---------- */

function desenharInicio() {
  const continuar = temProgresso();
  app.innerHTML = `
    <section class="inicio">
      <div class="selo">ANPOCS · 50 anos</div>
      <h1>Álbum de Figurinhas</h1>
      <p>Uma viagem pela memória visual dos Encontros da ANPOCS. Acerte as perguntas, ganhe pacotes e complete o álbum.</p>
      <button class="btn btn-verde" data-acao="entrar">${continuar ? "Continuar meu álbum" : "Começar"}</button>
      <button class="btn" data-acao="ajuda">Como funciona?</button>
      <p class="aviso-salvo">Seu progresso fica salvo neste aparelho.<br>
      Se abriu pelo QR code, prefira abrir no Chrome ou no Safari.</p>
    </section>`;
}

/* ---------- álbum ---------- */

function desenharAlbum(animacao = "") {
  const total = cartazes.length;
  const coladas = progresso.coladas.length;
  const porVez = paginasPorVez();
  tela.pagina = Math.max(0, Math.min(tela.pagina, totalPaginas() - 1));
  tela.pagina -= tela.pagina % porVez;

  const paginas = [];
  for (let p = tela.pagina; p < Math.min(tela.pagina + porVez, totalPaginas()); p++) {
    paginas.push(paginaHtml(p, animacao));
  }

  const falta = total - coladas - progresso.obtidas.length;
  const deck = progresso.obtidas.map(cartazPorId).filter(Boolean);
  const ultimaVisivel = Math.min(tela.pagina + porVez, totalPaginas());

  app.innerHTML = `
    <section class="album-tela">
      <header class="topo">
        <div>
          <div class="selo">ANPOCS · 50 anos</div>
          <h1>Álbum de Figurinhas</h1>
        </div>
        <button class="btn-ajuda" data-acao="ajuda" aria-label="Como funciona?">?</button>
      </header>

      <div class="progresso">
        <strong>${coladas} de ${total} coladas</strong>
        <div class="barra-progresso"><span style="width:${(100 * coladas) / total}%"></span></div>
        ${falta > 0 ? `<button class="btn btn-verde" data-acao="pergunta">Responder pergunta e ganhar pacote</button>` : ""}
      </div>

      <div class="livro ${porVez === 2 ? "duas" : ""}" id="livro">
        <div class="paginas-visiveis ${porVez === 2 ? "duas" : ""}">${paginas.join("")}</div>
      </div>

      <nav class="navegacao">
        <button class="btn" data-acao="anterior" ${tela.pagina === 0 ? "disabled" : ""} aria-label="Página anterior">←</button>
        <span class="contador">Página ${tela.pagina + 1}${porVez === 2 && ultimaVisivel > tela.pagina + 1 ? "–" + ultimaVisivel : ""} de ${totalPaginas()}</span>
        <button class="btn btn-principal" data-acao="proxima" ${ultimaVisivel >= totalPaginas() ? "disabled" : ""} aria-label="Próxima página">→</button>
      </nav>

      <section class="deck">
        <div class="deck-topo">
          <h2>Minhas figurinhas (${deck.length})</h2>
          ${deck.length ? `<button class="btn btn-principal" data-acao="colar-todas">Colar todas</button>` : ""}
        </div>
        ${deck.length
          ? `<p class="deck-ajuda">Toque numa figurinha para achar o lugar dela no álbum, ou use “Colar todas”.</p>
             <div class="deck-grade">${deck.map((c) => `
               <button class="deck-item ${tela.selecionada === c.id ? "selecionada" : ""}" data-acao="deck" data-id="${esc(c.id)}">
                 ${figurinhaHtml(c)}
               </button>`).join("")}</div>`
          : `<div class="deck-vazio">Nenhuma figurinha esperando para ser colada.</div>`}
      </section>

      <div class="rodape"><button class="btn-link" data-acao="recomecar">Recomeçar o álbum</button></div>
    </section>`;

  tela.recem.clear();
}

function tituloPagina(itens) {
  const anos = itens.filter((c) => !c.cromada && c.ano).map((c) => c.ano);
  if (!anos.length) return "Memória dos Encontros";
  return anos.length > 1 ? `${anos[0]} – ${anos[anos.length - 1]}` : anos[0];
}

function paginaHtml(p, animacao) {
  const itens = paginasAlbum[p] || [];
  const especiais = itens.length && itens.every((c) => c.cromada);
  const espacos = itens.map((c) => {
    if (progresso.coladas.includes(c.id)) {
      return `<div class="espaco espaco-colado ${tela.recem.has(c.id) ? "recem" : ""}" data-acao="espaco" data-id="${esc(c.id)}">${figurinhaHtml(c)}</div>`;
    }
    const obtida = progresso.obtidas.includes(c.id);
    const classe = tela.selecionada === c.id ? "espaco-alvo" : obtida ? "espaco-disponivel" : "";
    return `
      <div class="espaco ${classe}" data-acao="espaco" data-id="${esc(c.id)}">
        <div class="espaco-numero">${numeroFig(c.numero)}</div>
        <div class="espaco-ano">${esc(c.rotulo)}</div>
        <div class="espaco-texto">${obtida ? "Você tem esta! Toque para colar" : "Figurinha ainda não conquistada"}</div>
      </div>`;
  }).join("");

  return `
    <div class="pagina ${animacao}">
      <div class="pagina-topo">
        <div>
          <div class="pagina-rotulo">ANPOCS · 50 anos</div>
          <div class="pagina-titulo">${especiais ? "Figurinhas especiais" : tituloPagina(itens)}</div>
        </div>
        <div class="pagina-numero">${p + 1}</div>
      </div>
      <div class="grade-album">${espacos}</div>
      <div class="pagina-rodape">Álbum de Figurinhas · ANPOCS 50 anos</div>
    </div>`;
}

function irParaPagina(p, animacao = "") {
  tela.pagina = p;
  desenharAlbum(animacao);
}

function paginaDo(id) {
  return paginasAlbum.findIndex((pg) => pg.some((c) => c.id === id));
}

function mostrarEspaco(id) {
  tela.pagina = paginaDo(id);
  desenharAlbum();
  const alvo = app.querySelector(`.espaco[data-id="${CSS.escape(id)}"]`);
  if (alvo) alvo.scrollIntoView({ behavior: "smooth", block: "center" });
}

function colar(ids) {
  ids = ids.filter((id) => progresso.obtidas.includes(id));
  if (!ids.length) return;
  progresso.obtidas = progresso.obtidas.filter((id) => !ids.includes(id));
  progresso.coladas.push(...ids);
  tela.selecionada = null;
  ids.forEach((id) => tela.recem.add(id));
  salvar();
  mostrarEspaco(ids[0]);

  if (progresso.coladas.length === cartazes.length) {
    registrar("album/completou");
    setTimeout(abrirCompleto, 700);
  } else {
    toast(ids.length === 1
      ? `Figurinha ${cartazPorId(ids[0]).rotulo} colada!`
      : `${ids.length} figurinhas coladas!`);
  }
}

/* ---------- janelas ---------- */

function abrirJanela(html) {
  janela.innerHTML = `<div class="janela-caixa" role="dialog" aria-modal="true">${html}</div>`;
  janela.hidden = false;
  janela.scrollTop = 0;
}
function fecharJanela() {
  janela.hidden = true;
  janela.innerHTML = "";
}

const botaoVoltarAlbum = (rotulo = "Voltar ao álbum") =>
  `<button class="btn" data-acao="fechar">${esc(rotulo)}</button>`;

/* ---------- quiz ---------- */

let perguntaAtual = null;
let pacoteAtual = [];

function proximaPergunta() {
  return perguntas.find((p) => !progresso.respostas[p.id]) ||
         perguntas.find((p) => progresso.respostas[p.id] === "errada") || null;
}

function faltamFigurinhas() {
  const tem = new Set([...progresso.obtidas, ...progresso.coladas]);
  return cartazes.filter((c) => !tem.has(c.id));
}

function abrirPergunta() {
  if (!faltamFigurinhas().length) { toast("Você já ganhou todas as figurinhas!"); return; }
  perguntaAtual = proximaPergunta();
  if (!perguntaAtual) { toast("Não há mais perguntas por enquanto."); return; }
  const p = perguntaAtual;
  const numero = perguntas.indexOf(p) + 1;
  abrirJanela(`
    <div class="sobretitulo">Pergunta ${numero} de ${perguntas.length}</div>
    <h2 class="pergunta">${esc(p.enunciado)}</h2>
    <div class="alternativas">
      ${["a", "b", "c", "d"].map((l) =>
        `<button class="alternativa" data-acao="responder" data-letra="${l}">${esc(p[l])}</button>`).join("")}
    </div>
    <div class="linha-botoes">${botaoVoltarAlbum()}</div>`);
}

function responder(letra) {
  const p = perguntaAtual;
  const acertou = letra === p.correta;
  janela.querySelectorAll(".alternativa").forEach((b) => {
    b.disabled = true;
    if (b.dataset.letra === p.correta) b.classList.add("certa");
    else if (b.dataset.letra === letra) b.classList.add("errada");
  });

  progresso.respostas[p.id] = acertou ? "certa" : "errada";
  registrar(`album/pergunta/${p.id}/${acertou ? "acertou" : "errou"}`);
  if (acertou) {
    // o pacote já é guardado agora, para não se perder se a pessoa fechar a página
    pacoteAtual = faltamFigurinhas().slice(0, CONFIG.figurinhasPorPacote).map((c) => c.id);
    progresso.obtidas.push(...pacoteAtual);
  }
  salvar();

  const caixa = janela.querySelector(".linha-botoes");
  caixa.outerHTML = `
    <div class="retorno">
      <strong>${acertou ? "Você acertou!" : "Não foi dessa vez."}</strong>
      ${p.explicacao ? `<p>${esc(p.explicacao)}</p>` : ""}
    </div>
    <div class="linha-botoes">
      ${acertou
        ? `<button class="btn btn-verde" data-acao="pacote">Receber pacote com ${pacoteAtual.length} figurinhas</button>`
        : `<button class="btn btn-principal" data-acao="pergunta">Próxima pergunta</button>${botaoVoltarAlbum()}`}
    </div>`;
}

function pacoteHtml(rasgando) {
  return `
    <div class="sticker-pack ${rasgando ? "pack-tearing" : ""}" ${rasgando ? "" : 'data-acao="rasgar"'}>
      <div class="pack-spark ${rasgando ? "tear-spark" : ""}"></div>
      <div class="pack-top"></div>
      <div class="pack-rip-line"></div>
      <div class="pack-opening"></div>
      <div class="pack-fragment fragment-a"></div>
      <div class="pack-fragment fragment-b"></div>
      <div class="pack-fragment fragment-c"></div>
      <div class="pack-body"><strong>ANPOCS</strong><span>${pacoteAtual.length} figurinhas</span></div>
    </div>`;
}

function abrirPacote() {
  abrirJanela(`
    <div class="pack-screen">
      <div class="sobretitulo">Resposta correta</div>
      <h2 class="pergunta">Você ganhou um pacote!</h2>
      <p>Toque no pacote para abrir.</p>
      ${pacoteHtml(false)}
    </div>`);
}

function rasgarPacote() {
  abrirJanela(`
    <div class="pack-screen">
      <div class="sobretitulo">Abrindo o pacote</div>
      <h2 class="pergunta">Rasgando o lacre…</h2>
      ${pacoteHtml(true)}
    </div>`);
  setTimeout(revelarPacote, 1450);
}

function revelarPacote() {
  const itens = pacoteAtual.map(cartazPorId).filter(Boolean);
  abrirJanela(`
    <div class="pack-screen">
      <div class="sobretitulo">Pacote aberto</div>
      <h2 class="pergunta">Você encontrou ${itens.length} figurinhas!</h2>
      <div class="pack-reveal-grid">
        ${itens.map((c) => `<div class="pack-reveal-item">${figurinhaHtml(c)}</div>`).join("")}
      </div>
      <div class="linha-botoes">
        <button class="btn btn-principal" data-acao="colar-pacote">Colar no álbum</button>
        ${faltamFigurinhas().length ? `<button class="btn" data-acao="pergunta">Responder outra pergunta</button>` : ""}
        ${botaoVoltarAlbum("Guardar e ver o álbum")}
      </div>
    </div>`);
}

/* ---------- outras janelas ---------- */

function abrirDetalhe(id) {
  const c = cartazPorId(id);
  const campo = (rotulo, valor) => valor ? `<div class="campo"><span>${rotulo}</span><p>${esc(valor)}</p></div>` : "";
  const fonte = (rotulo, valor) => !valor ? "" : /^https?:\/\//i.test(valor)
    ? `<div class="campo"><span>${rotulo}</span><p><a href="${esc(valor)}" target="_blank" rel="noopener">Abrir fonte</a></p></div>`
    : campo(rotulo, valor);
  abrirJanela(`
    <div class="detalhe">
      ${figurinhaHtml(c)}
      <div>
        <div class="sobretitulo">Cartaz do Encontro</div>
        <h2>${esc([c.ano, c.edicao].filter(Boolean).join(" · ") || c.rotulo)}</h2>
        ${campo("Local", c.cidade)}
        ${campo("Tema", c.tema)}
        ${campo("Sobre esta edição", c.detalhes)}
        ${campo("Curiosidade", c.curiosidade)}
        ${fonte("Fonte das informações", c.fonteInfo)}
        ${fonte("Fonte da imagem", c.fonteImagem)}
      </div>
    </div>
    <div class="linha-botoes"><button class="btn btn-principal" data-acao="fechar">Fechar</button></div>`);
}

function abrirAjuda() {
  const passo = (n, t, d) => `<div class="passo"><b class="num">${n}</b><div><strong>${t}</strong><p>${d}</p></div></div>`;
  abrirJanela(`
    <h2 class="pergunta">Como funciona?</h2>
    ${passo(1, "Responda às perguntas", "Toque no botão verde “Responder pergunta”.")}
    ${passo(2, "Ganhe pacotes", `Cada acerto vale um pacote com ${CONFIG.figurinhasPorPacote} figurinhas. Errou? É só ir para a próxima pergunta.`)}
    ${passo(3, "Cole no álbum", "Toque numa figurinha e depois no espaço iluminado, ou use “Colar todas”.")}
    ${passo(4, "Conheça os cartazes", "Toque numa figurinha já colada para ver a história daquele Encontro.")}
    <div class="linha-botoes"><button class="btn btn-principal" data-acao="fechar">Entendi</button></div>`);
}

function abrirCompleto() {
  abrirJanela(`
    <div class="pack-screen">
      <div class="sobretitulo">Parabéns!</div>
      <h2 class="pergunta">Você completou o álbum dos 50 anos da ANPOCS!</h2>
      <p>São ${cartazes.length} figurinhas contando a história dos Encontros.</p>
      <div class="linha-botoes"><button class="btn btn-principal" data-acao="fechar">Ver meu álbum completo</button></div>
    </div>`);
}

function confirmarRecomeco() {
  abrirJanela(`
    <h2 class="pergunta">Recomeçar o álbum?</h2>
    <p>Todas as figurinhas e respostas deste aparelho serão apagadas.</p>
    <div class="linha-botoes">
      <button class="btn btn-perigo" data-acao="recomecar-sim">Sim, recomeçar</button>
      ${botaoVoltarAlbum("Não, voltar")}
    </div>`);
}

/* ---------- métricas (GoatCounter) ---------- */

function registrar(evento) {
  try {
    if (window.goatcounter && window.goatcounter.count) {
      window.goatcounter.count({ path: evento, title: evento, event: true });
    }
  } catch (e) {}
}

/* ---------- cliques ---------- */

document.addEventListener("click", (e) => {
  const alvo = e.target.closest("[data-acao]");
  if (!alvo) return;
  const id = alvo.dataset.id;

  switch (alvo.dataset.acao) {
    case "entrar":
      registrar(temProgresso() ? "album/continuou" : "album/comecou");
      desenharAlbum();
      break;
    case "ajuda": abrirAjuda(); break;
    case "fechar":
      fecharJanela();
      if (app.querySelector(".album-tela")) desenharAlbum(); else desenharInicio();
      break;
    case "pergunta": abrirPergunta(); break;
    case "responder": responder(alvo.dataset.letra); break;
    case "pacote": abrirPacote(); break;
    case "rasgar": rasgarPacote(); break;
    case "colar-pacote": fecharJanela(); colar([...pacoteAtual]); break;
    case "colar-todas": colar([...progresso.obtidas]); break;
    case "anterior": irParaPagina(tela.pagina - paginasPorVez(), "virando-tras"); break;
    case "proxima": irParaPagina(tela.pagina + paginasPorVez(), "virando-frente"); break;
    case "deck":
      tela.selecionada = id;
      mostrarEspaco(id);
      toast("Agora toque no espaço iluminado.");
      break;
    case "espaco":
      if (progresso.coladas.includes(id)) abrirDetalhe(id);
      else if (progresso.obtidas.includes(id)) colar([id]);
      else toast("Esta figurinha ainda não é sua. Responda perguntas para ganhar pacotes!");
      break;
    case "recomecar": confirmarRecomeco(); break;
    case "recomecar-sim":
      progresso = { obtidas: [], coladas: [], respostas: {} };
      salvar();
      tela.pagina = 0;
      tela.selecionada = null;
      fecharJanela();
      desenharInicio();
      break;
  }
});

/* ---------- arrastar o dedo para virar a página ---------- */

let toqueInicio = null;
document.addEventListener("pointerdown", (e) => {
  if (e.target.closest("#livro")) toqueInicio = { x: e.clientX, y: e.clientY };
});
document.addEventListener("pointerup", (e) => {
  if (!toqueInicio) return;
  const dx = e.clientX - toqueInicio.x;
  const dy = e.clientY - toqueInicio.y;
  toqueInicio = null;
  if (Math.abs(dx) < 60 || Math.abs(dy) > 60) return;
  const passo = paginasPorVez();
  if (dx < 0 && tela.pagina + passo < totalPaginas()) irParaPagina(tela.pagina + passo, "virando-frente");
  if (dx > 0 && tela.pagina > 0) irParaPagina(tela.pagina - passo, "virando-tras");
});

let larguraAnterior = paginasPorVez();
window.addEventListener("resize", () => {
  if (paginasPorVez() !== larguraAnterior && app.querySelector(".album-tela")) {
    larguraAnterior = paginasPorVez();
    desenharAlbum();
  }
});

/* ---------- início ---------- */

carregarPlanilha()
  .then(() => { lerProgresso(); desenharInicio(); })
  .catch((erro) => {
    console.error(erro);
    app.innerHTML = `<div class="erro"><strong>Não consegui ler a planilha <code>${esc(CONFIG.planilha)}</code>.</strong><br>
      Se abriu o arquivo direto do computador, use um servidor local — no R: <code>servr::httd("pasta_do_album")</code>.</div>`;
  });
