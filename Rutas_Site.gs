function renderUnifiedSitePage_(view) {
  const pages = {
    inicio:  { id: 'page-inicio', title: 'Global Transformation Hub' },
    kpis:    { id: 'page-kpis', title: 'KPIs EDC' },
    bloques: { id: 'page-endproduct', title: 'Bloques Funcionales' },
    ftes:    { id: 'page-ftes', title: 'FTEs y Presupuesto' },
    layout:  { id: 'page-catalogo-kpis', title: 'Layout KPIs' }
  };

  const selected = pages[String(view || '').toLowerCase()] || pages.inicio;
  const template = HtmlService.createTemplateFromFile('Index_Site');

  template.initialPage = selected.id;
  template.appUrl = getWebAppUrl_();
  template.initialDashboardSnapshot = 'null';

  if (selected.id === 'page-kpis' || selected.id === 'page-endproduct') {
    try {
      template.initialDashboardSnapshot = JSON.stringify(getSnapshotPayload_());
    } catch (error) {
      console.warn('No se pudo preparar la copia inicial del dashboard.', error);
    }
  }

  return template.evaluate()
    .setTitle(selected.title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
