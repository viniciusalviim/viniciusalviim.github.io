/* ============================================================
   EXPOSIÇÃO ANPOCS: 50 ANOS DE ENCONTROS
   O conteúdo vem da planilha conteudo_exposicao.xlsx.
   ============================================================ */

const CONFIG = {
  planilha: "conteudo_exposicao.xlsx",
  titulo: "ANPOCS: 50 anos de Encontros",
  selo: "Exposição · 50º Encontro Anual",
  cartaz: "imagens/cartaz_2026.jpg",
  pastaLogos: "imagens/logos/",
  secaoVideos: "O que é Anpocs para você?",
  segundosSemToque: 120, // depois disso aparece o "Ainda está aí?"
  segundosDeAviso: 20    // tempo do aviso antes de voltar ao início
};

// Seções de documentos (grade + página de leitura). Cada uma lê uma aba da planilha.
const ACERVOS = {
  imprensa: { titulo: "ANPOCS na imprensa", aba: "Imprensa", pasta: "imagens/imprensa/", fonte: "veiculo", icone: "❝" },
  cartas:   { titulo: "Cartas do acervo",   aba: "Cartas",   pasta: "imagens/cartas/",   fonte: "autor",   icone: "✉" }
};

const estado = {
  tela: "inicio",      // inicio | videos | video | lista | documento | acessibilidade
  acervo: "imprensa",  // qual seção de documentos está aberta
  telaAnterior: "inicio",
  item: null,
  dados: { videos: [], imprensa: [], cartas: [], logos: [] },
  contraste: false,
  letraGrande: false,
  voz: false,
  ampliado: false,
  decada: "todas"   // filtro por década
};

const app = document.getElementById("app");

/* ---------- utilidades ---------- */

const esc = (t) => String(t ?? "").replace(/[&<>"']/g, (c) => ({
  "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
}[c]));

const texto = (v) => String(v ?? "").trim();

function ativo(linha) {
  const v = texto(linha.ativo).toLowerCase();
  return v === "" || v === "sim" || v === "s";
}

function ordenar(lista) {
  return lista.sort((a, b) => (Number(a.ordem) || 9999) - (Number(b.ordem) || 9999));
}

function idYoutube(link) {
  const l = texto(link);
  if (!l) return "";
  const m = l.match(/(?:youtu\.be\/|v=|embed\/|shorts\/|live\/)([A-Za-z0-9_-]{11})/);
  if (m) return m[1];
  return /^[A-Za-z0-9_-]{11}$/.test(l) ? l : "";
}

/* ---------- planilha ---------- */

async function carregarPlanilha() {
  const resposta = await fetch(CONFIG.planilha + "?v=" + Date.now());
  if (!resposta.ok) throw new Error("Planilha não encontrada");
  const livro = XLSX.read(await resposta.arrayBuffer(), { type: "array" });
  const aba = (nome) => livro.Sheets[nome]
    ? XLSX.utils.sheet_to_json(livro.Sheets[nome], { defval: "" })
    : [];

  estado.dados.videos = ordenar(aba("Videos").filter((l) => texto(l.titulo) && ativo(l)))
    .map((l, i) => ({
      id: "v" + i,
      titulo: texto(l.titulo),
      descricao: texto(l.descricao),
      youtube: idYoutube(l.link_youtube)
    }));

  for (const [chave, acervo] of Object.entries(ACERVOS)) {
    estado.dados[chave] = ordenar(aba(acervo.aba).filter((l) => texto(l.titulo) && ativo(l)))
      .map((l, i) => ({
        id: chave + i,
        titulo: texto(l.titulo),
        fonte: texto(l[acervo.fonte]),
        data: texto(l.data),
        descricao: texto(l.descricao),
        arquivos: texto(l.arquivos).split(/[;,]/).map((a) => a.trim()).filter(Boolean)
      }));
  }

  estado.dados.logos = ordenar(aba("Logos").filter((l) => texto(l.arquivo) || texto(l.nome)))
    .map((l) => ({ arquivo: texto(l.arquivo), nome: texto(l.nome) }));
}

/* ---------- voz (leitura em voz alta) ---------- */

function falar(frase, forcar = false) {
  if (!("speechSynthesis" in window)) return;
  if (!estado.voz && !forcar) return;
  speechSynthesis.cancel();
  const msg = new SpeechSynthesisUtterance(frase);
  msg.lang = "pt-BR";
  msg.rate = 0.9;
  const vozPt = speechSynthesis.getVoices().find((v) => v.lang && v.lang.toLowerCase().startsWith("pt"));
  if (vozPt) msg.voice = vozPt;
  speechSynthesis.speak(msg);
}

function pararVoz() {
  if ("speechSynthesis" in window) speechSynthesis.cancel();
}

/* ---------- navegação ---------- */

function ir(tela, item = null) {
  pararVoz();
  destruirPlayer();
  estado.telaAnterior = estado.tela;
  estado.tela = tela;
  estado.item = item;
  estado.ampliado = false;
  desenhar();
  window.scrollTo(0, 0);
}

function voltar() {
  const mapa = { video: "videos", documento: "lista", acessibilidade: estado.telaAnterior };
  let destino = mapa[estado.tela] || "inicio";
  if (destino === "acessibilidade" || destino === estado.tela) destino = "inicio";
  const item = destino === "video" || destino === "documento" ? estado.item : null;
  ir(destino, item);
}

// Botão "voltar" do Android age como o "Voltar" da tela, sem sair do site
history.replaceState({ exposicao: 1 }, "");
history.pushState({ exposicao: 2 }, "");
window.addEventListener("popstate", () => {
  voltar();
  history.pushState({ exposicao: 2 }, "");
});

/* ---------- telas ---------- */

function barra(titulo) {
  return `
    <header class="barra">
      <button class="botao botao-principal" data-acao="voltar" data-falar="Voltar">← Voltar</button>
      <h1>${esc(titulo)}</h1>
      <button class="botao" data-acao="acessibilidade" data-falar="Acessibilidade" aria-label="Acessibilidade">Aa Acessibilidade</button>
    </header>`;
}

function telaInicio() {
  const logos = estado.dados.logos.map((l) => l.arquivo
    ? `<img src="${esc(CONFIG.pastaLogos + l.arquivo)}" alt="${esc(l.nome)}" data-reserva="${esc(l.nome)}">`
    : `<div class="logo-provisorio">${esc(l.nome)}</div>`).join("");

  return `
    <section class="inicio">
      <div class="inicio-cartaz">
        <img src="${esc(CONFIG.cartaz)}" alt="Cartaz do 50º Encontro Anual da ANPOCS" data-reserva="Cartaz 2026">
      </div>
      <div class="inicio-conteudo">
        <div class="inicio-selo">${esc(CONFIG.selo)}</div>
        <h1 class="inicio-titulo">${esc(CONFIG.titulo)}</h1>
        <div class="inicio-botoes">
          <button class="botao botao-principal botao-grande" data-acao="videos" data-falar="${esc(CONFIG.secaoVideos)}">
            <span class="icone" aria-hidden="true">▶</span> ${esc(CONFIG.secaoVideos)}
          </button>
          ${Object.entries(ACERVOS).map(([chave, a]) => `
          <button class="botao botao-principal botao-grande" data-acao="acervo" data-acervo="${chave}" data-falar="${esc(a.titulo)}">
            <span class="icone" aria-hidden="true">${a.icone}</span> ${esc(a.titulo)}
          </button>`).join("")}
          <button class="botao botao-grande" data-acao="acessibilidade" data-falar="Acessibilidade">
            <span class="icone" aria-hidden="true">Aa</span> Acessibilidade
          </button>
        </div>
      </div>
      ${logos ? `<footer class="logos">${logos}</footer>` : ""}
    </section>`;
}

function telaVideos() {
  const lista = estado.dados.videos;
  const cartoes = lista.map((v) => `
    <button class="cartao" data-acao="video" data-id="${v.id}" data-falar="${esc(v.titulo)}">
      <div class="cartao-imagem">
        ${v.youtube
          ? `<img src="https://i.ytimg.com/vi/${v.youtube}/hqdefault.jpg" alt=""><span class="selo-play" aria-hidden="true">▶</span>`
          : `<span>Vídeo em breve</span>`}
      </div>
      <div class="cartao-texto"><div class="cartao-titulo">${esc(v.titulo)}</div></div>
    </button>`).join("");

  return barra(CONFIG.secaoVideos) + `
    <div class="conteudo">
      ${lista.length ? `<div class="grade">${cartoes}</div>` : `<div class="vazio">Os vídeos serão publicados em breve.</div>`}
    </div>`;
}

function telaVideo() {
  const v = estado.item;
  return barra(CONFIG.secaoVideos) + `
    <div class="conteudo">
      <div class="pagina-video">
        <div class="player">
          ${v.youtube
            ? `<div id="player-youtube"></div>`
            : `<div class="player-aviso">Este vídeo será publicado em breve.</div>`}
        </div>
        <div class="legenda">
          <h2>${esc(v.titulo)}</h2>
          ${v.descricao ? `<p>${esc(v.descricao)}</p>` : ""}
          ${v.descricao ? `<div class="legenda-acoes"><button class="botao" data-acao="ouvir">Ouvir a descrição</button></div>` : ""}
        </div>
      </div>
    </div>`;
}

const decadaDe = (n) => {
  const ano = (String(n.data).match(/\d{4}/) || [])[0];
  return ano ? Math.floor(Number(ano) / 10) * 10 : null;
};

// "1984-10-24" ou "24/10/1984" → "24/10/1984"; "1984-10" → "10/1984"
function dataLegivel(d) {
  const m = String(d).match(/^(\d{4})-(\d{2})(?:-(\d{2}))?$/);
  if (!m) return d;
  return [m[3], m[2], m[1]].filter(Boolean).join("/");
}

function telaLista() {
  const acervo = ACERVOS[estado.acervo];
  const todas = estado.dados[estado.acervo];
  const decadas = [...new Set(todas.map(decadaDe).filter(Boolean))].sort((a, b) => a - b);
  if (estado.decada !== "todas" && !decadas.includes(Number(estado.decada))) estado.decada = "todas";
  const lista = estado.decada === "todas" ? todas : todas.filter((n) => decadaDe(n) === Number(estado.decada));

  const filtro = decadas.length > 1 ? `
    <div class="filtro-decadas" role="group" aria-label="Filtrar por década">
      ${["todas", ...decadas].map((d) => `
        <button class="botao ${String(estado.decada) === String(d) ? "botao-principal" : ""}"
                data-acao="decada" data-decada="${d}" aria-pressed="${String(estado.decada) === String(d)}"
                data-falar="${d === "todas" ? "Todas" : "Anos " + d}">
          ${d === "todas" ? "Todas" : "Anos " + d}
        </button>`).join("")}
    </div>` : "";

  // a grade usa a miniatura (pasta/miniaturas/arquivo); se não existir, usa a imagem inteira
  const cartoes = lista.map((n) => `
    <button class="cartao" data-acao="documento" data-id="${n.id}" data-falar="${esc(n.titulo)}">
      <div class="cartao-imagem">
        ${n.arquivos[0]
          ? `<img src="${esc(acervo.pasta + "miniaturas/" + n.arquivos[0])}" loading="lazy" alt=""
                  data-inteira="${esc(acervo.pasta + n.arquivos[0])}" data-reserva="Imagem em breve">`
          : `<span>Imagem em breve</span>`}
      </div>
      <div class="cartao-texto">
        <div class="cartao-titulo">${esc(n.titulo)}</div>
        <div class="cartao-sub">${esc([n.fonte, dataLegivel(n.data)].filter(Boolean).join(" · "))}</div>
      </div>
    </button>`).join("");

  return barra(acervo.titulo) + `
    <div class="conteudo">
      ${filtro}
      ${lista.length ? `<div class="grade grade-imprensa">${cartoes}</div>` : `<div class="vazio">Em breve.</div>`}
    </div>`;
}

function telaDocumento() {
  const acervo = ACERVOS[estado.acervo];
  const n = estado.item;
  const paginas = n.arquivos.map((a, i) =>
    `<img src="${esc(acervo.pasta + a)}" alt="${esc(n.titulo)} — página ${i + 1}">`).join("");
  return barra(acervo.titulo) + `
    <div class="conteudo noticia">
      <div class="legenda">
        <h2>${esc(n.titulo)}</h2>
        <div class="sub">${esc([n.fonte, dataLegivel(n.data)].filter(Boolean).join(" · "))}</div>
        ${n.descricao ? `<p>${esc(n.descricao)}</p>` : ""}
        <div class="legenda-acoes">
          ${paginas ? `<button class="botao botao-principal" data-acao="ampliar">${estado.ampliado ? "− Diminuir" : "+ Ampliar"}</button>` : ""}
          ${n.descricao ? `<button class="botao" data-acao="ouvir">Ouvir a descrição</button>` : ""}
        </div>
      </div>
      ${paginas ? `<div class="paginas ${estado.ampliado ? "ampliado" : ""}">${paginas}</div>` : ""}
    </div>`;
}

function telaAcessibilidade() {
  const opcao = (acao, rotulo, ligado) => `
    <button class="botao opcao" data-acao="${acao}" aria-pressed="${ligado}" data-falar="${esc(rotulo)}">
      <span>${esc(rotulo)}</span>
      <span class="estado">${ligado ? "Ligado" : "Desligado"}</span>
    </button>`;
  return barra("Acessibilidade") + `
    <div class="conteudo">
      <div class="opcoes">
        ${opcao("contraste", "Alto contraste", estado.contraste)}
        ${opcao("letra", "Letras maiores", estado.letraGrande)}
        ${opcao("voz", "Leitura em voz alta", estado.voz)}
      </div>
      <p class="dica">Use o fone de ouvido para ouvir os vídeos e a leitura em voz alta.
      Os vídeos têm legendas.</p>
    </div>`;
}

function desenhar() {
  document.documentElement.classList.toggle("contraste", estado.contraste);
  document.documentElement.classList.toggle("letra-grande", estado.letraGrande);

  const telas = {
    inicio: telaInicio, videos: telaVideos, video: telaVideo,
    lista: telaLista, documento: telaDocumento, acessibilidade: telaAcessibilidade
  };
  app.innerHTML = telas[estado.tela]();

  // imagem que não existe vira um espaço provisório
  app.querySelectorAll("img[data-reserva]").forEach((img) => {
    img.addEventListener("error", () => {
      if (img.dataset.inteira) { img.src = img.dataset.inteira; delete img.dataset.inteira; return; }
      const caixa = document.createElement("div");
      caixa.className = img.closest(".logos") ? "logo-provisorio"
        : img.closest(".inicio-cartaz") ? "cartaz-provisorio" : "";
      caixa.textContent = img.dataset.reserva;
      img.replaceWith(caixa);
    });
  });

  // notícia: cada página cresce até ocupar a largura ou a altura da tela, sem distorcer
  app.querySelectorAll(".paginas img").forEach((img) => {
    const ajustar = () => {
      const proporcao = img.naturalWidth / img.naturalHeight;
      img.style.width = `min(100%, calc((100svh - 8.6rem - 4px) * ${proporcao}))`;
    };
    if (img.complete && img.naturalWidth) ajustar(); else img.addEventListener("load", ajustar, { once: true });
  });

  if (estado.tela === "video" && estado.item.youtube) criarPlayer(estado.item.youtube);
}

/* ---------- métricas (GoatCounter) ---------- */

function registrar(evento) {
  try {
    if (window.goatcounter && window.goatcounter.count) {
      window.goatcounter.count({ path: evento, title: evento, event: true });
    }
  } catch (e) {}
}

// uma “visita” = alguém começa a usar o tablet a partir da tela inicial limpa
let visitaContada = false;

/* ---------- cliques ---------- */

app.addEventListener("click", (evento) => {
  const alvo = evento.target.closest("[data-acao]");
  if (!alvo) return;
  const acao = alvo.dataset.acao;

  if (alvo.dataset.falar) falar(alvo.dataset.falar);
  if (!visitaContada) { visitaContada = true; registrar("exposicao/visita"); }

  switch (acao) {
    case "voltar": voltar(); break;
    case "videos": registrar("exposicao/secao/videos"); ir("videos"); break;
    case "acervo":
      estado.acervo = alvo.dataset.acervo;
      estado.decada = "todas";
      registrar("exposicao/secao/" + estado.acervo);
      ir("lista");
      break;
    case "decada":
      registrar(`exposicao/${estado.acervo}/decada/` + alvo.dataset.decada);
      estado.decada = alvo.dataset.decada;
      desenhar();
      break;
    case "acessibilidade":
      if (estado.tela !== "acessibilidade") ir("acessibilidade", estado.item);
      break;
    case "video": {
      const v = estado.dados.videos.find((x) => x.id === alvo.dataset.id);
      registrar("exposicao/video/" + v.titulo);
      ir("video", v);
      falar(v.titulo + ". " + v.descricao);
      break;
    }
    case "documento": {
      const n = estado.dados[estado.acervo].find((x) => x.id === alvo.dataset.id);
      registrar(`exposicao/${estado.acervo}/` + n.titulo);
      ir("documento", n);
      falar(n.titulo + ". " + n.descricao);
      break;
    }
    case "ampliar":
      estado.ampliado = !estado.ampliado;
      app.querySelector(".paginas").classList.toggle("ampliado", estado.ampliado);
      alvo.textContent = estado.ampliado ? "− Diminuir" : "+ Ampliar";
      break;
    case "ouvir":
      falar(estado.item.descricao, true);
      break;
    case "contraste":
      estado.contraste = !estado.contraste;
      if (estado.contraste) registrar("exposicao/acessibilidade/alto-contraste");
      desenhar();
      break;
    case "letra":
      estado.letraGrande = !estado.letraGrande;
      if (estado.letraGrande) registrar("exposicao/acessibilidade/letras-maiores");
      desenhar();
      break;
    case "voz":
      estado.voz = !estado.voz;
      if (estado.voz) registrar("exposicao/acessibilidade/voz");
      desenhar();
      if (estado.voz) falar("Leitura em voz alta ligada.");
      else pararVoz();
      break;
  }
});

app.addEventListener("contextmenu", (e) => e.preventDefault());

/* ---------- YouTube ---------- */

let player = null;
let videoTocando = false;
let apiYoutube = null;

function carregarApiYoutube() {
  if (apiYoutube) return apiYoutube;
  apiYoutube = new Promise((resolve, reject) => {
    window.onYouTubeIframeAPIReady = resolve;
    const s = document.createElement("script");
    s.src = "https://www.youtube.com/iframe_api";
    s.onerror = reject;
    document.head.appendChild(s);
  });
  return apiYoutube;
}

async function criarPlayer(id) {
  try {
    await carregarApiYoutube();
  } catch {
    const caixa = document.getElementById("player-youtube");
    if (caixa) caixa.outerHTML = `<div class="player-aviso">Sem conexão com a internet. Chame o monitor da sala.</div>`;
    return;
  }
  if (!document.getElementById("player-youtube")) return; // saiu da tela antes de carregar
  player = new YT.Player("player-youtube", {
    videoId: id,
    playerVars: { rel: 0, playsinline: 1, hl: "pt-BR", cc_load_policy: 0 },
    events: {
      // os vídeos já têm legenda no próprio vídeo: desliga a legenda automática do YouTube
      onReady: (e) => desligarLegendaYoutube(e.target),
      onApiChange: (e) => desligarLegendaYoutube(e.target),
      onStateChange: (e) => {
        if (e.data === YT.PlayerState.PLAYING) desligarLegendaYoutube(e.target);
        videoTocando = e.data === YT.PlayerState.PLAYING || e.data === YT.PlayerState.BUFFERING;
        if (videoTocando) pararVoz();
        registrarAtividade();
      }
    }
  });
}

function desligarLegendaYoutube(p) {
  try {
    p.setOption("captions", "track", {});
    p.unloadModule("captions");
    p.unloadModule("cc");
  } catch (e) {}
}

function destruirPlayer() {
  if (player && player.destroy) player.destroy();
  player = null;
  videoTocando = false;
}

/* ---------- volta ao início por inatividade ---------- */

let ultimoToque = Date.now();
const aviso = document.getElementById("aviso");
const avisoContagem = document.getElementById("aviso-contagem");

function registrarAtividade() {
  ultimoToque = Date.now();
  if (!aviso.hidden) aviso.hidden = true;
}

["pointerdown", "keydown", "wheel", "touchstart"].forEach((ev) =>
  document.addEventListener(ev, registrarAtividade, { passive: true, capture: true }));
document.addEventListener("scroll", registrarAtividade, { passive: true, capture: true });
document.getElementById("aviso-continuar").addEventListener("click", registrarAtividade);

function telaJaEstaLimpa() {
  return estado.tela === "inicio" && !estado.contraste && !estado.letraGrande && !estado.voz;
}

function reiniciar() {
  aviso.hidden = true;
  estado.contraste = false;
  estado.letraGrande = false;
  estado.voz = false;
  estado.telaAnterior = "inicio";
  estado.decada = "todas";
  visitaContada = false;
  ir("inicio");
  ultimoToque = Date.now();
}

setInterval(() => {
  // vídeo tocando ou tela inicial já limpa: não conta o tempo
  if (videoTocando || telaJaEstaLimpa()) { ultimoToque = Date.now(); return; }

  const parado = (Date.now() - ultimoToque) / 1000;
  if (parado >= CONFIG.segundosSemToque + CONFIG.segundosDeAviso) {
    reiniciar();
  } else if (parado >= CONFIG.segundosSemToque) {
    aviso.hidden = false;
    avisoContagem.textContent = Math.ceil(CONFIG.segundosSemToque + CONFIG.segundosDeAviso - parado);
  }
}, 1000);

/* ---------- início ---------- */

carregarPlanilha()
  .then(desenhar)
  .catch((erro) => {
    console.error(erro);
    app.innerHTML = `<div class="erro">
      <strong>Não consegui ler a planilha <code>${esc(CONFIG.planilha)}</code>.</strong><br>
      Se você abriu o arquivo direto do computador (dando dois cliques),
      abra por um servidor local — no R: <code>servr::httd("pasta_da_exposicao")</code>.</div>`;
  });
