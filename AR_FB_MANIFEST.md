# Yve.01 — F&B Cost Control + AR Real Integration

## Completed in this session

### F&B Cost Control Module
- **fb_cost_control.py** — Core module (inventory, recipes, waste tracking, cost calculation)
- **tab_fb_dashboard.py** — Flask blueprint, SSE streaming, real-time KPI dashboard
- **datos-referencia/** — Reference data (recipes, inventory, waste, sales, KPIs by hotel)
- **Integration** — Tab in main dashboard (🍽️ F&B Cost button)

### AR Real Module — Corporate Group Reconciliation
- **ar_grupo_corporativo.py** — Core module (rooming reconciliation, invoice matching, 3-way matching)
- **tab_ar_grupo.py** — Flask blueprint, REST API for reconciliation
- **datos-referencia/hilton_abbvie_rooming.xlsx** — Real case: Abbvie Ovarian Cancer (87 rooms, 3 nights)
- **generador_reportes.py** — Export AR + F&B reports to Excel/PDF
- **Integration** — Tab in main dashboard (📋 AR Real button)

### Dashboard Integration
**dashboard.py** patches:
1. Import ar_bp + fb_bp
2. Register both blueprints
3. Add tab buttons (🍽️ F&B Cost, 📋 AR Real)
4. Add panel divs
5. Add switchTab handlers
6. Add async load functions (loadFBTab, loadARTab)
7. Add execution functions (runFB, runARPipeline)
8. Update role visibility

### Key Data Points (Hilton Abbvie)
- **Master ID**: 251527287
- **Grupo**: AbbVie Ovarian Cancer Educational Forum
- **Hotel**: Hilton Barcelona
- **Dates**: 2025-07-03 to 2025-07-06
- **Rooming**: 87 single rooms @ €210/night = €18,270 contracted
- **BEOs**: 3 (Setup + Meeting + Full Day) — €7,250 total setup
- **Invoice**: €1,081.35 (partial — individual invoices pending)
- **Variance**: €18,188.65 (under reconciliation)

### Deployment Checklist
- [ ] `git pull` to get all new modules
- [ ] `python dashboard.py` — test F&B + AR tabs locally
- [ ] Visit `/login` → admin/admin123
- [ ] Click 🍽️ F&B Cost → "Ejecutar" button
- [ ] Click 📋 AR Real → "Ejecutar AR" button
- [ ] Verify tabs load data + handle errors gracefully
- [ ] `git add . && git commit && git push` OR run `_git_push_ar.bat`
- [ ] Render auto-deploy (check build logs)
- [ ] Test https://yve01.onrender.com

### Files Structure
```
yve01/
├── dashboard.py (UPDATED — F&B + AR integration)
├── fb_cost_control.py (NEW)
├── tab_fb_dashboard.py (NEW)
├── ar_grupo_corporativo.py (NEW)
├── tab_ar_grupo.py (NEW)
├── generador_reportes.py (NEW)
├── setup_modules.py (NEW — helper)
├── datos-referencia/
│   ├── recetas.xlsx
│   ├── inventario.xlsx
│   ├── mermas.xlsx
│   ├── ventas_fb_diarias.xlsx
│   ├── hoteles.json
│   ├── kpis_hoteles.xlsx
│   └── hilton_abbvie_rooming.xlsx (NEW)
└── reportes/ (auto-created on first run)
```

### Known Issues / Next Steps
1. **Login still 302 on reload** — SECRET_KEY in Render needs manual fix (see render.yaml)
2. **F&B Ejecutar button** — needs actual invoice data from hotels
3. **AR Ejecutar AR button** — processes mock data, ready for real BEOs/rooming
4. **Reporting exports** — ready but not yet integrated into UI (future: add "Export" button)

### Quick Test Commands
```bash
# Local test
cd C:\Users\Jo\yve01
py dashboard.py

# Deploy
git add .
git commit -m "feat: F&B + AR complete"
git push
# → Render redeploys automatically
```

### Contact Points (Abbvie)
- **Organizer**: Louisa Worsfold (louisa.worsfold@stratacreate.com, +44 7923 992269)
- **Hotel**: Gemma Ràfols (events@hiltonbarcelona.com)
- **Billing**: AbbVie Inc., Strata Creative Communications Ltd (UK)

---
**Status**: ✅ Ready for production testing  
**Last updated**: 2026-06-05  
**Version**: 1.0 (F&B + AR Real integrated)
