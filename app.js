const TOTAL_SLIDES = 7;
const INTERACTIVE_SLIDE = 3;
const SWIPE_THRESHOLD = 50;
const DIRECTORY_FILE = "directorio-consolidado-juzgados-tabasco.xlsx";
const OFFICE_VIEWER_URL = "https://view.officeapps.live.com/op/embed.aspx?src=";
const OFFICIAL_SITES = window.OFFICIAL_SITES || [];

// Cada zona de las tablas abre los registros de su tipo en el directorio consolidado.
// La fila "Frontera" se muestra como "Centla" en el directorio oficial.
const MUNICIPIO_DIRECTORY = Object.freeze({
  "balancan-juzgados": "Balancán",
  "cardenas-juzgados": "Cárdenas",
  "frontera-juzgados": "Centla",
  "centro-juzgados": "Centro",
  "comalcalco-juzgados": "Comalcalco",
  "cunduacan-juzgados": "Cunduacán",
  "emiliano-zapata-juzgados": "Emiliano Zapata",
  "huimanguillo-juzgados": "Huimanguillo",
  "jalapa-juzgados": "Jalpa de Méndez",
  "jalapa-sierra-juzgados": "Jalapa",
  "jonuta-juzgados": "Jonuta",
  "macuspana-juzgados": "Macuspana",
  "nacajuca-juzgados": "Nacajuca",
  "paraiso-juzgados": "Paraíso",
  "tacotalpa-juzgados": "Tacotalpa",
  "teapa-juzgados": "Teapa",
  "tenosique-juzgados": "Tenosique",
  "nacajuca-salas": "Nacajuca",
  "centro-salas": "Centro",
  "huimanguillo-salas": "Huimanguillo",
  "cardenas-salas": "Cárdenas",
  "paraiso-salas": "Paraíso",
  "comalcalco-salas": "Comalcalco",
  "cunduacan-salas": "Cunduacán",
  "macuspana-salas": "Macuspana",
  "frontera-salas": "Centla",
  "emiliano-zapata-salas": "Emiliano Zapata",
  "jalapa-salas": "Jalapa"
});

const DIRECTORY_CATEGORY_LABELS = Object.freeze({
  juzgado: "Juzgados tradicionales",
  oralidad: "Salas de oralidad"
});

const tableCell = (id, name, x, y, width = 6.9, height = 2.2, category = "juzgado") => ({ id, name, x, y, width, height, category });

// Zonas clicables sobre los nombres de municipio en las dos tablas de la diapositiva 3.
const MUNICIPIOS = [
  // Salas de oralidad
  tableCell("nacajuca-salas", "Nacajuca", 4.6, 42.5, 5.2, 2.2, "oralidad"),
  tableCell("centro-salas", "Centro", 4.6, 45.2, 5.2, 2.2, "oralidad"),
  tableCell("huimanguillo-salas", "Huimanguillo", 18.5, 42.5, 6.9, 2.2, "oralidad"),
  tableCell("cardenas-salas", "Cárdenas", 18.5, 45.2, 6.9, 2.2, "oralidad"),
  tableCell("paraiso-salas", "Paraíso", 18.5, 47.9, 6.9, 2.2, "oralidad"),
  tableCell("comalcalco-salas", "Comalcalco", 18.5, 50.6, 6.9, 2.2, "oralidad"),
  tableCell("cunduacan-salas", "Cunduacán", 18.5, 53.1, 6.9, 2.2, "oralidad"),
  tableCell("macuspana-salas", "Macuspana", 32.9, 42.5, 6.9, 2.2, "oralidad"),
  tableCell("frontera-salas", "Frontera", 32.9, 45.2, 6.9, 2.2, "oralidad"),
  tableCell("emiliano-zapata-salas", "Emiliano Zapata", 47.4, 42.5, 6.9, 2.2, "oralidad"),
  tableCell("jalapa-salas", "Jalapa", 61.6, 42.5, 6.9, 2.2, "oralidad"),

  // Juzgados
  tableCell("centro-juzgados", "Centro", 4.6, 76.8, 5.2),
  tableCell("jalapa-juzgados", "Jalpa", 4.6, 79.2, 5.2),
  tableCell("nacajuca-juzgados", "Nacajuca", 4.6, 81.8, 5.2),
  tableCell("huimanguillo-juzgados", "Huimanguillo", 18.5, 76.8, 6.9, 1.65),
  tableCell("cardenas-juzgados", "Cárdenas", 18.5, 78.5, 6.9, 1.65),
  tableCell("paraiso-juzgados", "Paraíso", 18.5, 80.3, 6.9, 1.65),
  tableCell("comalcalco-juzgados", "Comalcalco", 18.5, 82.0, 6.9, 1.65),
  tableCell("cunduacan-juzgados", "Cunduacán", 18.5, 83.7, 6.9, 1.65),
  tableCell("macuspana-juzgados", "Macuspana", 32.9, 76.8),
  tableCell("jonuta-juzgados", "Jonuta", 32.9, 79.2),
  tableCell("frontera-juzgados", "Frontera", 32.9, 81.8),
  tableCell("emiliano-zapata-juzgados", "Emiliano Zapata", 47.4, 76.8),
  tableCell("tenosique-juzgados", "Tenosique", 47.4, 79.2),
  tableCell("balancan-juzgados", "Balancán", 47.4, 81.8),
  tableCell("jalapa-sierra-juzgados", "Jalapa", 61.6, 76.8),
  tableCell("tacotalpa-juzgados", "Tacotalpa", 61.6, 79.2),
  tableCell("teapa-juzgados", "Teapa", 61.6, 81.8)
];

const state = { currentSlide: 1 };
const directoryState = { map: null, initialized: false };
const el = {
  slide: document.querySelector("#slide"),
  slideBackground: document.querySelector("#slideBackground"),
  hotspots: document.querySelector("#hotspots"),
  template: document.querySelector("#hotspotTemplate"),
  mapDialog: document.querySelector("#mapDialog"),
  mapDialogCategory: document.querySelector("#mapDialogCategory"),
  mapDialogTitle: document.querySelector("#mapDialogTitle"),
  mapCanvas: document.querySelector("#mapCanvas"),
  directoryDialog: document.querySelector("#directoryDialog"),
  directoryViewer: document.querySelector("#directoryViewer"),
  directoryTabs: Array.from(document.querySelectorAll("[data-directory-view]")),
  directorySitesView: document.querySelector("#officialSitesView"),
  directoryExcelView: document.querySelector("#directoryExcelView"),
  officialSitesMap: document.querySelector("#officialSitesMap")
};

function renderSlide() {
  const interactive = state.currentSlide === INTERACTIVE_SLIDE;
  el.slideBackground.style.backgroundImage = `url("diapositiva-${state.currentSlide}.png")`;
  el.hotspots.hidden = !interactive;
}

function renderHotspots() {
  for (const municipio of MUNICIPIOS) {
    const button = el.template.content.firstElementChild.cloneNode(true);
    button.style.left = `${municipio.x}%`;
    button.style.top = `${municipio.y}%`;
    button.style.width = `${municipio.width}%`;
    button.style.height = `${municipio.height}%`;
    const categoryLabel = DIRECTORY_CATEGORY_LABELS[municipio.category];
    button.title = `Abrir ${categoryLabel.toLowerCase()} de ${municipio.name}`;
    button.setAttribute("aria-label", `Abrir ${categoryLabel.toLowerCase()} de ${municipio.name}`);
    button.querySelector(".sr-only").textContent = `Abrir ${categoryLabel.toLowerCase()} de ${municipio.name}`;
    button.addEventListener("click", () => openMunicipalityDirectory(municipio));
    el.hotspots.append(button);
  }
}

function changeSlide(change) {
  state.currentSlide = Math.max(1, Math.min(TOTAL_SLIDES, state.currentSlide + change));
  renderSlide();
}

let gesture = null;

function beginGesture(event) {
  if (event.button !== 0 || event.target.closest(".hotspot")) return;

  gesture = {
    pointerId: event.pointerId,
    startX: event.clientX,
    startY: event.clientY
  };

  el.slide.setPointerCapture?.(event.pointerId);
}

function endGesture(event) {
  if (!gesture || event.pointerId !== gesture.pointerId) return;

  const deltaX = event.clientX - gesture.startX;
  const deltaY = event.clientY - gesture.startY;
  gesture = null;
  el.slide.releasePointerCapture?.(event.pointerId);

  const isHorizontalSwipe = Math.abs(deltaX) >= SWIPE_THRESHOLD && Math.abs(deltaX) > Math.abs(deltaY);

  if (isHorizontalSwipe) {
    changeSlide(deltaX < 0 ? 1 : -1);
    return;
  }

  if (Math.abs(deltaX) < 8 && Math.abs(deltaY) < 8) {
    const bounds = el.slide.getBoundingClientRect();
    const direction = event.clientX < bounds.left + bounds.width / 2 ? -1 : 1;
    changeSlide(direction);
  }
}

function cancelGesture(event) {
  if (!gesture || event.pointerId !== gesture.pointerId) return;

  gesture = null;
  el.slide.releasePointerCapture?.(event.pointerId);
}

function openMunicipalityDirectory(municipio) {
  const directoryMunicipality = MUNICIPIO_DIRECTORY[municipio.id] || municipio.name;
  const categoryLabel = DIRECTORY_CATEGORY_LABELS[municipio.category];
  const records = window.COURTS_BY_CATEGORY?.[municipio.category]?.[directoryMunicipality] || [];

  el.mapDialogTitle.textContent = directoryMunicipality;
  el.mapDialogCategory.textContent = categoryLabel;
  el.mapCanvas.replaceChildren();

  if (records.length) {
    renderCourtDirectory(directoryMunicipality, records, categoryLabel);
  } else {
    renderDirectoryPlaceholder(directoryMunicipality, categoryLabel);
  }
  el.mapDialog.showModal();
}

function renderCourtDirectory(municipality, records, categoryLabel) {
  const directory = document.createElement("section");
  directory.className = "court-directory";

  const introduction = document.createElement("p");
  introduction.className = "court-directory__intro";
  introduction.textContent = `${records.length} ${records.length === 1 ? "registro disponible" : "registros disponibles"} de ${categoryLabel.toLowerCase()}. La ubicación de cada sede se muestra junto a su ficha.`;
  directory.append(introduction);

  const list = document.createElement("div");
  list.className = "court-directory__list";
  records.forEach((record) => list.append(createCourtCard(municipality, record)));
  directory.append(list);
  el.mapCanvas.append(directory);
}

function createCourtCard(municipality, record) {
  const card = document.createElement("article");
  card.className = "court-card";

  const heading = document.createElement("h3");
  heading.textContent = record.court;
  const details = document.createElement("dl");
  appendCourtDetail(details, "Juez", record.judge);
  appendCourtDetail(details, "Teléfono", record.extension && record.extension !== "0" ? `${record.phone} · Ext. ${record.extension}` : record.phone);
  appendCourtDetail(details, "Dirección", record.address);
  appendCourtDetail(details, "Código postal", record.postalCode && record.postalCode !== "0" ? record.postalCode : "No disponible");

  const actions = document.createElement("div");
  actions.className = "court-card__actions";
  if (record.source) {
    const source = document.createElement("a");
    source.className = "court-card__source";
    source.href = record.source;
    source.target = "_blank";
    source.rel = "noopener";
    source.textContent = "Fuente oficial";
    actions.append(source);
  }

  const map = document.createElement("div");
  map.className = "court-card__map";
  const query = [record.court, record.address, municipality, "Tabasco, México"].filter(Boolean).join(", ");
  const frame = document.createElement("iframe");
  frame.className = "court-card__map-frame";
  frame.loading = "lazy";
  frame.allowFullscreen = true;
  frame.referrerPolicy = "no-referrer-when-downgrade";
  frame.title = `Ubicación de ${record.judge || record.court}`;
  frame.src = `https://www.google.com/maps?q=${encodeURIComponent(query)}&output=embed`;
  map.append(frame);

  card.append(heading, details);
  if (actions.childElementCount) card.append(actions);
  card.append(map);
  return card;
}

function appendCourtDetail(details, label, value) {
  if (!value) return;
  const term = document.createElement("dt");
  term.textContent = label;
  const description = document.createElement("dd");
  description.textContent = value;
  details.append(term, description);
}

function renderDirectoryPlaceholder(municipio, categoryLabel) {
  const message = document.createElement("div");
  message.className = "map-placeholder";
  const heading = document.createElement("strong");
  heading.textContent = `${categoryLabel} de ${municipio}`;
  const text = document.createElement("span");
  text.textContent = "No se encontraron registros de este tipo para el municipio en el directorio consolidado.";
  message.append(heading, text);
  el.mapCanvas.append(message);
}

function openDirectoryDialog() {
  el.directoryDialog.showModal();
  setDirectoryView("sites");
}

function setDirectoryView(view) {
  const showingSites = view === "sites";
  el.directorySitesView.hidden = !showingSites;
  el.directoryExcelView.hidden = showingSites;
  el.directoryTabs.forEach((tab) => {
    const active = tab.dataset.directoryView === view;
    tab.classList.toggle("is-active", active);
    tab.setAttribute("aria-selected", String(active));
  });

  if (showingSites) {
    requestAnimationFrame(renderOfficialSitesMap);
    return;
  }

  if (!el.directoryViewer.src) {
    const fileUrl = new URL(DIRECTORY_FILE, window.location.href).href;
    el.directoryViewer.src = `${OFFICE_VIEWER_URL}${encodeURIComponent(fileUrl)}`;
  }
}

function renderOfficialSitesMap() {
  if (directoryState.initialized) {
    directoryState.map.invalidateSize();
    return;
  }

  if (!window.L || !OFFICIAL_SITES.length) {
    el.officialSitesMap.classList.add("official-sites-map--unavailable");
    el.officialSitesMap.textContent = "No fue posible cargar las sedes oficiales.";
    return;
  }

  const map = window.L.map(el.officialSitesMap, { scrollWheelZoom: false });
  window.L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 19,
    attribution: "© OpenStreetMap contributors"
  }).addTo(map);

  const bounds = [];
  OFFICIAL_SITES.forEach((site) => {
    const marker = window.L.marker([site.lat, site.lng]);
    const popup = document.createElement("div");
    const name = document.createElement("strong");
    name.textContent = site.name;
    popup.append(name);
    marker.bindPopup(popup).bindTooltip(site.name, { direction: "top", offset: [0, -28] }).addTo(map);
    bounds.push([site.lat, site.lng]);
  });

  map.fitBounds(bounds, { padding: [28, 28] });
  directoryState.map = map;
  directoryState.initialized = true;
}

document.querySelectorAll("[data-close-dialog]").forEach((button) => {
  button.addEventListener("click", () => document.querySelector(`#${button.dataset.closeDialog}`).close());
});
document.querySelector("[data-open-directory]").addEventListener("click", openDirectoryDialog);
el.directoryTabs.forEach((tab) => {
  tab.addEventListener("click", () => setDirectoryView(tab.dataset.directoryView));
});
document.addEventListener("keydown", (event) => {
  if (document.querySelector("dialog[open]")) return;
  if (event.key === "ArrowLeft") changeSlide(-1);
  if (event.key === "ArrowRight") changeSlide(1);
});

el.slide.addEventListener("pointerdown", beginGesture);
el.slide.addEventListener("pointerup", endGesture);
el.slide.addEventListener("pointercancel", cancelGesture);

renderHotspots();
renderSlide();
