/* yve-icons v2 - la tabla de iconos de Yve, en un solo sitio.
 *
 * La usan el dashboard y la pantalla de facturas por aprobar. Antes vivia
 * dentro del literal HTML de dashboard.py y la segunda pantalla no la tenia:
 * alli todos los simbolos salian como emoji del sistema operativo, en color,
 * al lado de los SVG monocromos del resto de la app.
 *
 * Anadir un icono = una entrada en P (el trazado) y una en MAP (que emoji lo
 * dispara). Los emoji que NO estan en MAP se quedan tal cual, a proposito:
 * las banderas del selector de idioma, los puntos de estado de color y las
 * flechas tipograficas.
 */
(function(){
if (window.icon) return;   // ya cargado
var P={
 zap:'<polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/>',
 camera:'<path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z"/><circle cx="12" cy="13" r="4"/>',
 dots:'<circle cx="5" cy="12" r="1.6" fill="currentColor" stroke="none"/><circle cx="12" cy="12" r="1.6" fill="currentColor" stroke="none"/><circle cx="19" cy="12" r="1.6" fill="currentColor" stroke="none"/>',
 inbox:'<polyline points="22 12 16 12 14 15 10 15 8 12 2 12"/><path d="M5.45 5.11 2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z"/>',
 pack:'<path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/>',
 chart:'<line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/>',
 bank:'<line x1="3" y1="22" x2="21" y2="22"/><line x1="6" y1="18" x2="6" y2="11"/><line x1="10" y1="18" x2="10" y2="11"/><line x1="14" y1="18" x2="14" y2="11"/><line x1="18" y1="18" x2="18" y2="11"/><polygon points="12 2 20 7 4 7"/>',
 bell:'<path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/>',
 food:'<path d="M3 2v7c0 1.1.9 2 2 2h4a2 2 0 0 0 2-2V2"/><path d="M7 2v20"/><path d="M21 15V2a5 5 0 0 0-5 5v6c0 1.1.9 2 2 2h3zm0 0v7"/>',
 build:'<path d="M6 22V4a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v18Z"/><path d="M6 12H4a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2h2"/><path d="M18 9h2a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2h-2"/><path d="M10 6h4"/><path d="M10 10h4"/><path d="M10 14h4"/><path d="M10 18h4"/>',
 hotel:'<path d="M2 4v16"/><path d="M2 8h18a2 2 0 0 1 2 2v10"/><path d="M2 17h20"/><path d="M6 8v9"/>',
 globe:'<circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/>',
 file:'<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/>',
 trend:'<polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/>',
 target:'<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/>',
 clip:'<path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><rect x="8" y="2" width="8" height="4" rx="1"/>',
 user:'<path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/>',
 users:'<path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/>',
 key:'<path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4"/>',
 money:'<rect x="2" y="6" width="20" height="12" rx="2"/><circle cx="12" cy="12" r="2.5"/><path d="M6 12h.01"/><path d="M18 12h.01"/>',
 wrench:'<path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/>',
 refresh:'<polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10"/><path d="M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/>',
 out:'<path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/>',
 palette:'<path d="M12 22a10 10 0 1 1 10-10c0 1.66-1.34 3-3 3h-2a2 2 0 0 0-2 2c0 .5.2 1 .54 1.36.34.36.46.9.4 1.4A2 2 0 0 1 14 22z"/><circle cx="13.5" cy="6.5" r=".9" fill="currentColor" stroke="none"/><circle cx="17.5" cy="10.5" r=".9" fill="currentColor" stroke="none"/><circle cx="8.5" cy="7.5" r=".9" fill="currentColor" stroke="none"/><circle cx="6.5" cy="12" r=".9" fill="currentColor" stroke="none"/>',
 phone:'<rect x="5" y="2" width="14" height="20" rx="2"/><line x1="12" y1="18" x2="12.01" y2="18"/>',
 search:'<circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>',
 save:'<path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/>',
 folder:'<path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>',
 up:'<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>',
 down:'<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>',
 img:'<rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/>',
 sunrise:'<path d="M12 2v6"/><path d="m4.93 10.93 1.41 1.41"/><path d="M2 18h2"/><path d="M20 18h2"/><path d="m19.07 10.93-1.41 1.41"/><path d="M22 22H2"/><path d="m8 6 4-4 4 4"/><path d="M16 18a4 4 0 0 0-8 0"/>',
 trophy:'<path d="M6 9H4.5a2.5 2.5 0 0 1 0-5H6"/><path d="M18 9h1.5a2.5 2.5 0 0 0 0-5H18"/><path d="M4 22h16"/><path d="M10 14.66V17c0 .55-.47.98-.97 1.21C7.85 18.75 7 20.24 7 22"/><path d="M14 14.66V17c0 .55.47.98.97 1.21C16.15 18.75 17 20.24 17 22"/><path d="M18 2H6v7a6 6 0 0 0 12 0V2z"/>',
 moon:'<path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/>',
 sun:'<circle cx="12" cy="12" r="5"/><path d="M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42"/>',
 pin:'<path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/>',
 mail:'<rect x="2" y="4" width="20" height="16" rx="2"/><polyline points="22 6 12 13 2 6"/>',
 card:'<rect x="1" y="4" width="22" height="16" rx="2"/><line x1="1" y1="10" x2="23" y2="10"/>',
 spark:'<path d="M12 3l1.9 5.7a2 2 0 0 0 1.3 1.3L21 12l-5.8 1.9a2 2 0 0 0-1.3 1.3L12 21l-1.9-5.8a2 2 0 0 0-1.3-1.3L3 12l5.8-1.9a2 2 0 0 0 1.3-1.3z"/>',
 chat:'<path d="M21 11.5a8.5 8.5 0 0 1-8.5 8.5 8.4 8.4 0 0 1-3.8-.9L3 21l1.9-5.7a8.4 8.4 0 0 1-.9-3.8 8.5 8.5 0 0 1 8.5-8.5 8.5 8.5 0 0 1 8.5 8.5z"/>',
 mask:'<path d="M4 7c2.5-1.3 5.2-2 8-2s5.5.7 8 2v4a8 8 0 0 1-16 0Z"/><circle cx="9" cy="10" r="1" fill="currentColor" stroke="none"/><circle cx="15" cy="10" r="1" fill="currentColor" stroke="none"/><path d="M9 14c1 .8 2 1.2 3 1.2s2-.4 3-1.2"/>',
 x:'<line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>',
 check:'<polyline points="20 6 9 17 4 12"/>',
 warn:'<path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>',
 info:'<circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/>',
 clock:'<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>',
 trash:'<polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/>',
 edit:'<path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.12 2.12 0 0 1 3 3L12 15l-4 1 1-4z"/>',
 book:'<path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/>',
 receipt:'<path d="M5 2h14v20l-2.5-1.6L14 22l-2-1.6L10 22l-2.5-1.6L5 22z"/><line x1="9" y1="8" x2="15" y2="8"/><line x1="9" y1="12" x2="15" y2="12"/>',
 bot:'<rect x="4" y="8" width="16" height="12" rx="2"/><path d="M12 8V5"/><circle cx="12" cy="3.5" r="1.2"/><line x1="9" y1="13" x2="9" y2="14.5"/><line x1="15" y1="13" x2="15" y2="14.5"/><path d="M2 13v3"/><path d="M22 13v3"/>',
 lock:'<rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>',
 keyboard:'<rect x="2" y="6" width="20" height="12" rx="2"/><path d="M6 10h.01M10 10h.01M14 10h.01M18 10h.01M6 14h.01M18 14h.01"/><line x1="9.5" y1="14" x2="14.5" y2="14"/>',
 gear:'<circle cx="12" cy="12" r="3.2"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/>',
 case:'<rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/>',
 scales:'<path d="M12 3v18"/><path d="M5 7h14"/><path d="M8 21h8"/><path d="M5 7 2 14a3 3 0 0 0 6 0z"/><path d="M19 7l-3 7a3 3 0 0 0 6 0z"/>',
 help:'<circle cx="12" cy="12" r="10"/><path d="M9.1 9a3 3 0 0 1 5.82 1c0 2-3 3-3 3"/><line x1="12" y1="17" x2="12.01" y2="17"/>',
 wave:'<path d="M18 11V6.5a1.75 1.75 0 0 0-3.5 0V11"/><path d="M14.5 10.5V4.75a1.75 1.75 0 0 0-3.5 0V11"/><path d="M11 11V6.25a1.75 1.75 0 0 0-3.5 0V14"/><path d="M18 8.5a1.75 1.75 0 0 1 3.5 0V14a8 8 0 0 1-8 8h-1a8 8 0 0 1-8-8 1.75 1.75 0 0 1 3.5 0"/>',
 star:'<polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>',
 eye:'<path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/>',
 calendar:'<rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>',
 bulb:'<path d="M9 18h6"/><path d="M10 22h4"/><path d="M15.09 14c.18-.98.65-1.74 1.41-2.5A4.65 4.65 0 0 0 18 8 6 6 0 0 0 6 8c0 1 .23 2.23 1.5 3.5A4.61 4.61 0 0 1 8.91 14"/>',
 map:'<polygon points="1 6 8 3 16 6 23 3 23 18 16 21 8 18 1 21 1 6"/><line x1="8" y1="3" x2="8" y2="18"/><line x1="16" y1="6" x2="16" y2="21"/>',
 flask:'<path d="M9 3h6"/><path d="M10 3v6l-5.5 9A2 2 0 0 0 6.2 21h11.6a2 2 0 0 0 1.7-3L14 9V3"/><path d="M7.5 15h9"/>',
 rocket:'<path d="M4.5 16.5c-1.5 1.26-2 5-2 5s3.74-.5 5-2c.71-.84.7-2.13-.09-2.91a2.18 2.18 0 0 0-2.91 0z"/><path d="m12 15-3-3a22 22 0 0 1 2-3.95A12.88 12.88 0 0 1 22 2c0 2.72-.78 7.5-6 11a22.35 22.35 0 0 1-4 2z"/><path d="M9 12H4s.55-3.03 2-4c1.62-1.08 5 0 5 0"/><path d="M12 15v5s3.03-.55 4-2c1.08-1.62 0-5 0-5"/>',
 home:'<path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/>',
 tag:'<path d="M20.59 13.41 12 22l-9-9V3h10l7.59 7.59a2 2 0 0 1 0 2.82z"/><line x1="7" y1="7" x2="7.01" y2="7"/>',
 plane:'<path d="M17.8 19.2 16 11l3.5-3.5a2.12 2.12 0 0 0-3-3L13 8 4.8 6.2a1 1 0 0 0-.9 1.7l5.6 3.4-2.3 2.3-2.6-.5a1 1 0 0 0-.9 1.6l2.6 2.6a1 1 0 0 0 1.6-.9l-.5-2.6 2.3-2.3 3.4 5.6a1 1 0 0 0 1.7-.9z"/>',
 link:'<path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>',
 back:'<line x1="19" y1="12" x2="5" y2="12"/><polyline points="12 19 5 12 12 5"/>'
};
window.icon=function(n,cls){var p=P[n];if(!p)return '';
 return '<svg class="yvi'+(cls?' '+cls:'')+'" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'+p+'</svg>';};
var MAP={
 '\u26a1':'zap','\ud83d\udd0c':'zap','\ud83d\udcf8':'camera','\ud83d\udcf7':'camera','\u22ef':'dots',
 '\u22ee':'dots','\ud83d\udce5':'inbox','\ud83d\udced':'inbox','\ud83d\udcec':'inbox','\ud83d\udce8':'inbox',
 '\ud83d\udce9':'inbox','\ud83d\udce6':'pack','\ud83d\ude9a':'pack','\ud83d\uded2':'pack','\ud83d\udcca':'chart',
 '\ud83d\udcc9':'chart','\ud83c\udfe6':'bank','\ud83d\udd14':'bell','\ud83d\udd15':'bell','\ud83c\udf7d':'food',
 '\ud83c\udf74':'food','\ud83e\udd58':'food','\ud83c\udfe2':'build','\ud83c\udfec':'build','\ud83c\udfed':'build',
 '\ud83c\udfe8':'hotel','\ud83c\udfe9':'hotel','\ud83d\udecf':'hotel','\ud83d\udece':'hotel','\ud83c\udf0d':'globe',
 '\ud83c\udf0e':'globe','\ud83c\udf0f':'globe','\ud83c\udf10':'globe','\ud83d\uddfa':'map','\ud83d\udcc4':'file',
 '\ud83d\udcc3':'file','\ud83d\udcd1':'file','\ud83d\udcf0':'file','\ud83d\udcc8':'trend','\ud83c\udfaf':'target',
 '\ud83d\udccb':'clip','\ud83d\udcce':'clip','\ud83d\udc64':'user','\ud83d\ude4b':'user','\ud83d\udc65':'users',
 '\ud83e\udd1d':'users','\ud83d\udd11':'key','\ud83d\udddd':'key','\ud83d\udcb0':'money','\ud83d\udcb5':'money',
 '\ud83d\udcb6':'money','\ud83d\udcb4':'money','\ud83d\udcb7':'money','\ud83d\udee0':'wrench','\ud83d\udd27':'wrench',
 '\ud83d\udd28':'wrench','\u21bb':'refresh','\u21ba':'refresh','\ud83d\udd04':'refresh','\ud83d\udd01':'refresh',
 '\ud83d\udd03':'refresh','\u21a9':'out','\ud83d\udeaa':'out','\ud83c\udfa8':'palette','\ud83d\udcf1':'phone',
 '\ud83d\udcf2':'phone','\ud83d\udcde':'phone','\u260e':'phone','\ud83d\udd0d':'search','\ud83d\udd0e':'search',
 '\ud83d\udcbe':'save','\ud83d\udcbf':'save','\ud83d\udcc2':'folder','\ud83d\udcc1':'folder','\ud83d\uddc2':'folder',
 '\ud83d\uddc3':'folder','\ud83d\uddc4':'folder','\ud83d\udce4':'up','\u2b06':'up','\u2b07':'down',
 '\ud83d\uddbc':'img','\ud83c\udf9e':'img','\ud83c\udf05':'sunrise','\ud83c\udf04':'sunrise','\ud83c\udfc6':'trophy',
 '\ud83e\udd47':'trophy','\ud83c\udf19':'moon','\u2600':'sun','\ud83c\udf1e':'sun','\ud83d\udccd':'pin',
 '\ud83d\udccc':'pin','\ud83d\udce7':'mail','\u2709':'mail','\ud83d\udcee':'mail','\ud83d\udcb3':'card',
 '\ud83d\udd2e':'spark','\u2728':'spark','\ud83d\udcab':'spark','\ud83c\udd95':'spark','\ud83c\udf89':'spark',
 '\ud83c\udf8a':'spark','\ud83d\udcac':'chat','\ud83d\udde8':'chat','\ud83c\udfad':'mask','\u2715':'x',
 '\u274c':'x','\u2717':'x','\u2718':'x','\u274e':'x','\u2713':'check',
 '\u2705':'check','\u2714':'check','\u2611':'check','\u26a0':'warn','\ud83d\udea8':'warn',
 '\u2757':'warn','\u2755':'warn','\u203c':'warn','\u2139':'info','\u24d8':'info',
 '\u23f3':'clock','\u231b':'clock','\u23f0':'clock','\u23f1':'clock','\ud83d\uddd1':'trash',
 '\ud83d\udcdd':'edit','\u270d':'edit','\u270f':'edit','\ud83d\udd8a':'edit','\ud83d\udcd6':'book',
 '\ud83d\udcd2':'book','\ud83d\udcd5':'book','\ud83d\udcd7':'book','\ud83d\udcd8':'book','\ud83d\udcd9':'book',
 '\ud83d\udcd3':'book','\ud83d\udcda':'book','\ud83e\uddfe':'receipt','\ud83e\udd16':'bot','\ud83d\udd12':'lock',
 '\ud83d\udd10':'lock','\ud83d\udd0f':'lock','\ud83d\udee1':'lock','\u2328':'keyboard','\u2699':'gear',
 '\ud83c\udf9b':'gear','\ud83d\udcbc':'case','\u2696':'scales','\u2753':'help','\u2754':'help',
 '\ud83d\udc4b':'wave','\u2b50':'star','\ud83c\udf1f':'star','\ud83d\udc41':'eye','\ud83d\udc40':'eye',
 '\ud83d\udcc5':'calendar','\ud83d\uddd3':'calendar','\ud83d\udcc6':'calendar','\ud83d\udca1':'bulb','\ud83e\uddea':'flask',
 '\u2697':'flask','\ud83d\udd2c':'flask','\ud83d\ude80':'rocket','\ud83c\udfe0':'home','\ud83c\udfe1':'home',
 '\ud83c\udff7':'tag','\ud83c\udfab':'tag','\ud83c\udf9f':'tag','\u2708':'plane','\ud83d\udd17':'link',
 '\u2190':'back'};
var RX=new RegExp('('+Object.keys(MAP).map(function(c){return c.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');}).join('|')+')\uFE0F?','g');
var SKIP={SCRIPT:1,STYLE:1,TEXTAREA:1,INPUT:1,SELECT:1,OPTION:1,TITLE:1};
function iconizeIn(root){
 try{
  var r=(root&&root.nodeType)?root:document.body; if(!r) return;
  var w=document.createTreeWalker(r,NodeFilter.SHOW_TEXT,{acceptNode:function(n){
    var p=n.parentNode; if(!p||SKIP[p.nodeName]||(p.closest&&p.closest('svg'))) return NodeFilter.FILTER_REJECT;
    RX.lastIndex=0; return RX.test(n.nodeValue)?NodeFilter.FILTER_ACCEPT:NodeFilter.FILTER_SKIP;
  }},false);
  var nodes=[]; while(w.nextNode()) nodes.push(w.currentNode);
  nodes.forEach(function(n){
    var txt=n.nodeValue, frag=document.createDocumentFragment(), last=0, m; RX.lastIndex=0;
    while((m=RX.exec(txt))!==null){
      if(m.index>last) frag.appendChild(document.createTextNode(txt.slice(last,m.index)));
      var sp=document.createElement('span'); sp.innerHTML=window.icon(MAP[m[1]]);
      if(sp.firstChild) frag.appendChild(sp.firstChild); last=RX.lastIndex;
    }
    if(last<txt.length) frag.appendChild(document.createTextNode(txt.slice(last)));
    n.parentNode.replaceChild(frag,n);
  });
 }catch(e){}
}
window.iconizeIn=iconizeIn;
window._iconizeAll=function(){ if(typeof _saveOriginals==='function'){try{_saveOriginals();}catch(e){}} iconizeIn(document.body); };
var _icoT=null;
function boot(){
  window._iconizeAll();
  try{
    new MutationObserver(function(){ clearTimeout(_icoT); _icoT=setTimeout(function(){ iconizeIn(document.body); },120); })
      .observe(document.body,{childList:true,subtree:true,characterData:true});
  }catch(e){}
}
if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',function(){ setTimeout(boot,150); });
else setTimeout(boot,150);
})();
