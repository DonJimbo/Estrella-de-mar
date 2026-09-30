function getDashboardPayloadForClient(action) {
  const requestedAction = String(action || 'getSnapshot').trim();

  if (requestedAction === 'getSnapshot' || requestedAction === 'getData') {
    return getSnapshotPayload_();
  }

  if (requestedAction === 'getCacheStatus') {
    return { status: 'success', data: getCacheStatus_() };
  }

  if (requestedAction === 'refreshCache' || requestedAction === 'refreshData') {
    return refreshDashboardCache_();
  }

  if (requestedAction === 'getEquiposEDC') {
    return { status: 'success', data: getEquiposEDC_() };
  }

  if (requestedAction === 'getMisSolicitudes') {
    return { status: 'success', data: getMisSolicitudes_() };
  }

  throw new Error('Acción no reconocida: ' + requestedAction);
}
