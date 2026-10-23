const C='torta-runner-v29';
const FILES=["./", "index.html", "config.js", "manifest.json", "icon-192.png", "icon-512.png", "favicon-32.png", "assets/cinthia.webp", "assets/deisy.webp", "assets/dina.webp", "assets/dina_avatar.webp", "assets/dina_corre.webp", "assets/fiesta.mp4", "assets/fiesta.webm", "assets/fiesta_poster.webp", "assets/frutilla.webp", "assets/leandro.webp", "assets/nora.webp", "assets/kampito.webp", "assets/kampito_tropical.webp", "assets/kampito_verde.webp", "assets/nubes.webp", "assets/pancito.webp", "assets/sofi.webp", "assets/sofi_hoodie.webp", "assets/texto.webp", "assets/torta.webp", "assets/vela.mp4", "assets/vela.webm", "assets/vela_poster.webp"];
// cache:'reload' baja cada archivo de nuevo (sin copias viejas del navegador) al instalar una versión nueva
self.addEventListener('install',e=>{e.waitUntil(caches.open(C).then(c=>c.addAll(FILES.map(f=>new Request(f,{cache:'reload'})))));self.skipWaiting()});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==C).map(k=>caches.delete(k)))));self.clients.claim()});
self.addEventListener('fetch',e=>{
  const u=new URL(e.request.url);
  if(e.request.method!=='GET'||u.origin!==location.origin) return; // el ranking (Google) va siempre a la red
  if(u.pathname.endsWith('config.js')){ // la configuración: primero la red, así un cambio se ve enseguida
    e.respondWith(fetch(e.request).then(res=>{const cp=res.clone();caches.open(C).then(c=>c.put(e.request,cp));return res}).catch(()=>caches.match(e.request)));return;
  }
  e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request).then(res=>{const cp=res.clone();caches.open(C).then(c=>c.put(e.request,cp));return res})));
});
