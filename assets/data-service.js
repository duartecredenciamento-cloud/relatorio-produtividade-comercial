window.DataService = {
  async loadDashboard(filters = {}) {
    const query = new URLSearchParams(Object.entries(filters).filter(([, value]) => value !== "" && value != null));
    const response = await fetch(`/api/dashboard${query.size ? `?${query}` : ""}`, {
      headers: { Accept: "application/json" }, cache: "no-store"
    });
    if (!response.ok) throw new Error(`Não foi possível atualizar o painel (${response.status}).`);
    return response.json();
  },
  async refresh() {
    const response = await fetch("/api/refresh", { method: "POST", cache: "no-store" });
    if (!response.ok) throw new Error("Não foi possível solicitar a atualização.");
    return response.json();
  },
  async setMarker(recordId, marked) {
    const response = await fetch("/api/marker", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ record_id: recordId, marked }) });
    if (!response.ok) throw new Error("Não foi possível salvar a revisão da marcação.");
    return response.json();
  },
  async setGoal(service, target, period = "manual") {
    const response = await fetch("/api/goals", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ service, target, period }) });
    if (!response.ok) throw new Error("Não foi possível salvar a meta.");
    return response.json();
  }
};
