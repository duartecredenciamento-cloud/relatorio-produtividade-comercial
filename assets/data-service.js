window.DataService = {
  async request(url, options = {}) {
    const response = await fetch(url, { credentials: "same-origin", cache: "no-store", ...options });
    if (response.status === 401) { window.location.replace("/login"); throw new Error("Sessão encerrada. Entre novamente."); }
    return response;
  },
  async health() {
    const response = await this.request("/api/health", { headers: { Accept: "application/json" } });
    if (!response.ok) throw new Error("Não foi possível verificar o estado da leitura.");
    return response.json();
  },
  async loadDashboard(filters = {}) {
    const query = new URLSearchParams(Object.entries(filters).filter(([, value]) => value !== "" && value != null));
    const response = await this.request(`/api/dashboard${query.size ? `?${query}` : ""}`, { headers: { Accept: "application/json" } });
    if (!response.ok) throw new Error(`Não foi possível atualizar o painel (${response.status}).`);
    return response.json();
  },
  async refresh() {
    const response = await this.request("/api/refresh", { method: "POST" });
    if (!response.ok) throw new Error("Não foi possível solicitar a atualização.");
    return response.json();
  },
  async setMarker(recordId, marked) {
    const response = await this.request("/api/marker", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ record_id: recordId, marked }) });
    if (!response.ok) throw new Error("Não foi possível salvar a revisão da marcação.");
    return response.json();
  },
  async setGoal(service, target, period = "manual") {
    const response = await this.request("/api/goals", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ service, target, period }) });
    if (!response.ok) throw new Error("Não foi possível salvar a meta.");
    return response.json();
  },
  async logout() {
    return this.request("/api/logout", { method: "POST" });
  }
};
