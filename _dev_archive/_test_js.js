// Mock browser globals for syntax check
const document = { getElementById: () => ({set textContent(v){}, getContext: () => ({}), classList: {add(){},remove(){},toggle(){}}, appendChild(){}, scrollTop:0, scrollHeight:0, get style(){return {}}, set style(v){}, innerHTML:'', value:'', focus(){}, get disabled(){return false}, set disabled(v){}}), querySelectorAll: () => [], createElement: () => ({set className(v){}, set textContent(v){}, appendChild(){}, classList:{add(){},remove(){},toggle(){}}, get style(){return {}}, set innerHTML(v){}, set value(v){}}), };
const fetch = () => Promise.resolve({json:()=>Promise.resolve({}), ok:true});
const EventSource = function(){this.onmessage=null;this.onerror=null;this.close=()=>{}};
const Chart = function(){};
Chart.prototype.destroy = ()=>{};
const setInterval = () => {};
const setTimeout = () => {};
const Intl = {NumberFormat: function(){return {format:()=>''}}};
const JSON = {stringify:()=>'',parse:()=>({})};
const console = {log(){},warn(){},error(){}};
const Date = function(){return {toLocaleDateString:()=>'',toLocaleTimeString:()=>''}};

// ── Actual JS from dashboard.py below ──
