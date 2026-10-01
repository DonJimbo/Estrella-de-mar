function getDashboardPayloadForClient(action) {
  const requestedAction = String(action || 'getSnapshot').trim();

  if (requestedAction === 'getSnapshot' || requestedAction === 'getData' ||
      requestedAction === 'refreshCache' || requestedAction === 'refreshData') {
    // Sin caché: cualquiera de estas acciones exporta el libro en el momento.
    return getSnapshotPayload_();
  }

  if (requestedAction === 'getEquiposEDC') {
    return { status: 'success', data: getEquiposEDC_() };
  }

  if (requestedAction === 'getMisSolicitudes') {
    return { status: 'success', data: getMisSolicitudes_() };
  }

  throw new Error('Acción no reconocida: ' + requestedAction);
}
