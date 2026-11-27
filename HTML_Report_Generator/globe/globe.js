/* ── globe: verbatim port of C:\Users\User\Desktop\Global Purchase Report
 * Claude\3D Earth Dashboard\build\src\globe.js (git-twin of the shipped GPR
 * earth engine), read at build time by MTC HTML_Report_Generator\build_report.js
 * and inlined into the Dashboard sheet. EXACTLY FOUR seams vs the source:
 *   1. (this header) provenance comment restated for MTC.
 *   2. `readJson` prepended: in GPR it comes from core.js:39; MTC's sheet does
 *      not port core.js, so the one-line helper lives here (globe.json seam).
 *   3. loadLand island id: 'gpr-land' → 'mtc-land' (MTC island names are
 *      mtc-three / mtc-land / mtc-globe-data; everything else about the island
 *      contract is unchanged).
 *   4. MTC CITY extension below the table + an onHover param for the readout
 *      panel (MTC UI spec: hover country/USD/share; GPR only had click).
 * Original header kept below for reference. ────────────────────────────────
 * Faithful copy with three changes for the single-file build:
 *   1. `import * as THREE`  → THREE arrives as the first argument (blob-module
 *      loaded on demand, see loadThree() in core-adjacent boot code).
 *   2. geo.load() fetches land-110m.json → the TopoJSON is embedded in the
 *      page (#gpr-land) and decoded synchronously; file:// has no fetch.
 *   3. Only the tok/rgb helpers from fx.js are kept (the animation canvases
 *      the template shipped are unused by this screen). */

/* MTC SEAM 2: the only core.js piece this port needs (core.js:39). */
const readJson = (id) => JSON.parse(document.getElementById(id).textContent);

const RM = matchMedia('(prefers-reduced-motion: reduce)');

/* token reader with theme-invalidated cache — from fx.js */
const _tokCache = { cache: {}, gen: 0 };
const tokBump = () => { _tokCache.cache = {}; _tokCache.gen++ };
addEventListener('kit:theme', tokBump);
new MutationObserver(tokBump).observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });
const tok = (name) => _tokCache.cache[name] ?? (_tokCache.cache[name] = getComputedStyle(document.documentElement).getPropertyValue(name).trim());

function rgbTok(v) {
  if (!v) return [128, 128, 128];
  v = v.trim();
  if (v[0] === '#') {
    const h = v.length === 4 ? v.slice(1).split('').map((c) => c + c).join('') : v.slice(1);
    return [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16));
  }
  const n = (v.match(/[\d.]+/g) || []).map(Number);
  if (/^color\(|^oklch|^lab/.test(v)) {
    const c = document.createElement('canvas').getContext('2d');
    c.fillStyle = v; return rgbTok(c.fillStyle);
  }
  return n.length >= 3 ? n.slice(0, 3) : [128, 128, 128];
}

/* ── geo (sync loader) ─────────────────────────────────────────────────── */
let _topo = null, _rings = null;
const _rasters = new Map();
const loadLand = () => { if (_topo) return _topo; _topo = readJson('mtc-land'); return _topo }; /* MTC SEAM 3: island id */

function ringsOf(topo) {
  if (_rings) return _rings;
  const { scale: [sx, sy], translate: [tx, ty] } = topo.transform;
  const arc = (i) => {
    const rev = i < 0, a = topo.arcs[rev ? ~i : i];
    let x = 0, y = 0;
    const pts = a.map(([dx, dy]) => [(x += dx) * sx + tx, (y += dy) * sy + ty]);
    return rev ? pts.reverse() : pts;
  };
  const ring = (idx) => idx.reduce((acc, i) => acc.concat(acc.length ? arc(i).slice(1) : arc(i)), []);
  _rings = [];
  for (const g of topo.objects.land.geometries) {
    const polys = g.type === 'Polygon' ? [g.arcs] : g.arcs;
    for (const p of polys) for (const r of p) _rings.push(ring(r));
  }
  return _rings;
}

function raster(w = 2048) {
  if (_rasters.has(w)) return _rasters.get(w);
  const h = w / 2;
  const cv = document.createElement('canvas');
  cv.width = w; cv.height = h;
  const c = cv.getContext('2d', { willReadFrequently: true });
  c.fillStyle = '#fff';
  c.beginPath();
  for (const r of ringsOf(loadLand())) {
    r.forEach(([lon, lat], i) => {
      const x = (lon + 180) / 360 * w, y = (90 - lat) / 180 * h;
      i ? c.lineTo(x, y) : c.moveTo(x, y);
    });
    c.closePath();
  }
  c.fill('evenodd');
  const d = c.getImageData(0, 0, w, h).data;
  const hit = (lon, lat) => {
    const x = Math.min(w - 1, Math.max(0, (lon + 180) / 360 * w | 0));
    const y = Math.min(h - 1, Math.max(0, (90 - lat) / 180 * h | 0));
    return d[(y * w + x) * 4 + 3] > 90;
  };
  const out = { w, h, data: d, hit };
  _rasters.set(w, out);
  return out;
}

function dotsOf({ step = 1.5, latMax = 83 } = {}) {
  const r = raster(2048), out = [];
  for (let lat = -latMax; lat <= latMax; lat += step) {
    const ring = Math.max(1, Math.round(360 / (step / Math.cos(lat * Math.PI / 180))));
    for (let i = 0; i < ring; i++) {
      const lon = -180 + i * 360 / ring;
      if (r.hit(lon, lat)) out.push([lon, lat]);
    }
  }
  return out;
}

/* city table (kept whole — future routes may use it). [lon,lat] */
const CITY = {
  '北京': [116.4, 39.9], '上海': [121.5, 31.2], '深圳': [114.1, 22.5], '广州': [113.3, 23.1],
  '成都': [104.1, 30.7], '西安': [108.9, 34.3], '杭州': [120.2, 30.3], '武汉': [114.3, 30.6],
  '沈阳': [123.4, 41.8], '哈尔滨': [126.6, 45.8], '南京': [118.8, 32.1], '重庆': [106.5, 29.6],
  '天津': [117.2, 39.1], '长沙': [113.0, 28.2], '苏州': [120.6, 31.3], '昆明': [102.7, 25.0],
  '香港': [114.2, 22.3], '新加坡': [103.8, 1.3], '东京': [139.7, 35.7], '首尔': [127.0, 37.6],
  '曼谷': [100.5, 13.8], '孟买': [72.9, 19.1], '迪拜': [55.3, 25.3], '莫斯科': [37.6, 55.8],
  '法兰克福': [8.7, 50.1], '伦敦': [-0.1, 51.5], '巴黎': [2.4, 48.9], '阿姆斯特丹': [4.9, 52.4],
  '罗马': [12.5, 41.9],
  '纽约': [-74.0, 40.7], '洛杉矶': [-118.2, 34.1], '芝加哥': [-87.6, 41.9], '圣保罗': [-46.6, -23.6],
  '利马': [-77.0, -12.0],
  '悉尼': [151.2, -33.9], '约翰内斯堡': [28.0, -26.2], '开罗': [31.2, 30.0], '墨西哥城': [-99.1, 19.4],
};

/* MTC SEAM 4a: CITY extension — every key the MTC COUNTRY_COORDS map (defined
 * in build_report.js and baked into the mtc-globe-data island as each entry's
 * `h`) needs but GPR's table lacks. The original table above is untouched.
 * 利雅得 = Riyadh (46.71 E, 24.71 N) — user-confirmed hub for ledger value
 * 'MIDDLE-EAST'; the six sourcing-side hubs are the BRIEF addendum's explicit
 * choices (Shanghai/Taipei already partially present — only missing keys added). */
Object.assign(CITY, {
  '利雅得': [46.71, 24.71],   // Riyadh — MIDDLE-EAST hub (user decision)
  '柏林': [13.4, 52.5],       // Berlin — GERMANY
  '多伦多': [-79.4, 43.7],    // Toronto — CANADA (economic hub)
  '波哥大': [-74.1, 4.7],     // Bogotá — COLOMBIA
  '哥本哈根': [12.6, 55.7],   // Copenhagen — DENMARK
  '台北': [121.5, 25.0],      // Taipei — TAIWAN (origin side)
  '河内': [105.85, 21.03],    // Hanoi — VIETNAM (origin side)
  '米兰': [9.19, 45.46],      // Milan — ITALY (origin side, addendum-specified)
  '雅加达': [106.85, -6.21],  // Jakarta — INDONESIA (origin side)
  '达卡': [90.41, 23.81],     // Dhaka — BANGLADESH (origin side)
});

/* ── the globe itself ──────────────────────────────────────────────────── */
function makeGlobe(THREE, host, { routes = [], markers = [], step = 1.5, spin = .03, tilt = .32, dayLength = 90, interactive = true, onPick = null, onHover = null } = {}) {
  loadLand();

  const R = 1;
  const rad = Math.PI / 180;
  const HEAD = 14;
  const TAIL = .16;
  const toVec = ([lon, lat], r = R) => new THREE.Vector3(
    r * Math.cos(lat * rad) * Math.cos(lon * rad),
    r * Math.sin(lat * rad),
    -r * Math.cos(lat * rad) * Math.sin(lon * rad),
  );
  const col = (name) => new THREE.Color(...rgbTok(tok(name)).map((v) => v / 255));
  const num = (name, dflt) => { const v = parseFloat(tok(name)); return Number.isFinite(v) ? v : dflt };

  const cv = document.createElement('canvas');
  cv.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;display:block;' +
                     (interactive ? 'cursor:grab' : 'pointer-events:none');
  if (getComputedStyle(host).position === 'static') host.style.position = 'relative';
  host.appendChild(cv);

  const renderer = new THREE.WebGLRenderer({ canvas: cv, antialias: true, alpha: true });
  renderer.setPixelRatio(Math.min(devicePixelRatio || 1, 2));
  const scene = new THREE.Scene();
  const camera = new THREE.PerspectiveCamera(32, 1, .1, 100);
  camera.position.set(0, 0, 3.7);

  const world = new THREE.Group();
  world.rotation.x = tilt;
  scene.add(world);

  /* The sun lives in world space so it does NOT rotate with the earth — the
     ground turns underneath a fixed light, which is what makes a terminator. */
  const sun = new THREE.Vector3(1, .22, .45).normalize();

  const ocean = new THREE.Mesh(
    new THREE.SphereGeometry(R * .995, 96, 64),
    new THREE.ShaderMaterial({
      uniforms: {
        uDeep: { value: col('--globe-deep') }, uSea: { value: col('--globe-sea') },
        uSun: { value: sun }, uNight: { value: num('--globe-night-k', .32) },
      },
      vertexShader: `varying vec3 vN; varying vec3 vV; varying vec3 vW;
        void main(){
          vN = normalize(normalMatrix * normal);
          vW = normalize(mat3(modelMatrix) * normal);
          vec4 mv = modelViewMatrix * vec4(position,1.); vV = normalize(-mv.xyz);
          gl_Position = projectionMatrix * mv;
        }`,
      fragmentShader: `uniform vec3 uDeep, uSea, uSun; uniform float uNight;
        varying vec3 vN; varying vec3 vV; varying vec3 vW;
        void main(){
          float day = smoothstep(-0.18, 0.32, dot(vW, uSun));
          float fres = pow(1.0 - max(dot(vN, vV), 0.0), 2.6);
          vec3 c = mix(uDeep * uNight, mix(uDeep, uSea, 0.55), day);
          c += uSea * fres * (0.25 + day * 0.5);
          gl_FragColor = vec4(c, 1.0);
        }`,
    }));
  world.add(ocean);

  const air = new THREE.Mesh(
    new THREE.SphereGeometry(R * 1.035, 64, 44),
    new THREE.ShaderMaterial({
      uniforms: { uCol: { value: col('--globe-air') }, uSun: { value: sun }, uA: { value: num('--globe-air-a', .9) } },
      vertexShader: `varying vec3 vN; varying vec3 vV; varying vec3 vW;
        void main(){
          vN = normalize(normalMatrix * normal);
          vW = normalize(mat3(modelMatrix) * normal);
          vec4 mv = modelViewMatrix * vec4(position,1.); vV = normalize(-mv.xyz);
          gl_Position = projectionMatrix * mv;
        }`,
      fragmentShader: `uniform vec3 uCol; uniform vec3 uSun; uniform float uA;
        varying vec3 vN; varying vec3 vV; varying vec3 vW;
        void main(){
          float rim = pow(max(1.0 - abs(dot(vN, vV)), 0.0), 5.5);
          float lit = smoothstep(-0.55, 0.45, dot(vW, uSun));
          gl_FragColor = vec4(uCol, rim * (0.22 + lit * 0.78) * uA);
        }`,
      side: THREE.BackSide, blending: THREE.AdditiveBlending, transparent: true, depthWrite: false,
    }));
  world.add(air);

  /* ~14k land points, one draw call; a stable pseudo-random weight lights only
     about a third of them on the night side, latitudinally damped → "cities". */
  const pts = dotsOf({ step });
  const pos = new Float32Array(pts.length * 3);
  const cityW = new Float32Array(pts.length);
  pts.forEach((p, i) => {
    toVec(p, R * 1.002).toArray(pos, i * 3);
    const h = Math.abs(Math.sin(p[0] * 12.9898 + p[1] * 78.233) * 43758.5453) % 1;
    cityW[i] = h < .34 * (1 - Math.min(1, Math.abs(p[1]) / 72) * .75) ? 1 : 0;
  });
  const landGeo = new THREE.BufferGeometry();
  landGeo.setAttribute('position', new THREE.BufferAttribute(pos, 3));
  landGeo.setAttribute('aCity', new THREE.BufferAttribute(cityW, 1));
  const landMat = new THREE.ShaderMaterial({
    uniforms: {
      uDay: { value: col('--globe-land') }, uDark: { value: col('--globe-landnight') },
      uCity: { value: col('--globe-city') }, uSun: { value: sun },
      uSize: { value: 2.4 * Math.min(devicePixelRatio || 1, 2) },
    },
    vertexShader: `uniform float uSize; uniform vec3 uSun;
      attribute float aCity;
      varying float vFace; varying float vDay; varying float vCity;
      void main(){
        vec4 mv = modelViewMatrix * vec4(position, 1.0);
        vFace = smoothstep(-0.12, 0.5, normalize(mat3(modelViewMatrix) * position).z);
        vDay  = smoothstep(-0.14, 0.3, dot(normalize(mat3(modelMatrix) * position), uSun));
        vCity = aCity;
        gl_PointSize = uSize * (2.6 / -mv.z) * (1.0 + aCity * (1.0 - vDay) * 0.5);
        gl_Position = projectionMatrix * mv;
      }`,
    fragmentShader: `uniform vec3 uDay, uDark, uCity;
      varying float vFace; varying float vDay; varying float vCity;
      void main(){
        float d = length(gl_PointCoord - 0.5);
        if (d > 0.5) discard;
        float night = 1.0 - vDay;
        vec3 c = mix(uDark, uDay, vDay);
        float glow = vCity * night;
        c = mix(c, uCity, glow * 0.9);
        float a = (1.0 - d * 1.25) * vFace * (0.55 + vDay * 0.45 + glow * 0.6);
        gl_FragColor = vec4(c, a);
      }`,
    transparent: true, depthWrite: false,
  });
  world.add(new THREE.Points(landGeo, landMat));

  /* coastlines — the definition pass; antimeridian seams discarded. */
  const coast = (() => {
    const v = [];
    for (const ring of ringsOf(loadLand())) {
      for (let i = 1; i < ring.length; i++) {
        if (Math.abs(ring[i][0] - ring[i - 1][0]) > 170) continue;
        const a = toVec(ring[i - 1], R * 1.006), b = toVec(ring[i], R * 1.006);
        v.push(a.x, a.y, a.z, b.x, b.y, b.z);
      }
    }
    const g = new THREE.BufferGeometry();
    g.setAttribute('position', new THREE.BufferAttribute(new Float32Array(v), 3));
    return new THREE.LineSegments(g, new THREE.ShaderMaterial({
      uniforms: { uCol: { value: col('--globe-coast') }, uSun: { value: sun }, uA: { value: num('--globe-coast-a', .85) } },
      vertexShader: `uniform vec3 uSun; varying float vFace; varying float vDay;
        void main(){
          vec4 mv = modelViewMatrix * vec4(position, 1.0);
          vFace = smoothstep(-0.05, 0.42, normalize(mat3(modelViewMatrix) * position).z);
          vDay  = smoothstep(-0.2, 0.35, dot(normalize(mat3(modelMatrix) * position), uSun));
          gl_Position = projectionMatrix * mv;
        }`,
      fragmentShader: `uniform vec3 uCol; uniform float uA; varying float vFace; varying float vDay;
        void main(){ gl_FragColor = vec4(uCol, vFace * (0.3 + vDay * 0.7) * uA); }`,
      transparent: true, depthWrite: false,
    }));
  })();
  world.add(coast);

  const grat = new THREE.LineSegments(
    (() => {
      const v = [];
      const push = (a, b) => { v.push(a.x, a.y, a.z, b.x, b.y, b.z) };
      for (let lat = -60; lat <= 60; lat += 30)
        for (let i = 0; i < 120; i++)
          push(toVec([i * 3 - 180, lat], R * 1.003), toVec([(i + 1) * 3 - 180, lat], R * 1.003));
      for (let lon = -180; lon < 180; lon += 30)
        for (let i = 0; i < 60; i++)
          push(toVec([lon, i * 3 - 90], R * 1.003), toVec([lon, (i + 1) * 3 - 90], R * 1.003));
      const g = new THREE.BufferGeometry();
      g.setAttribute('position', new THREE.BufferAttribute(new Float32Array(v), 3));
      return g;
    })(),
    new THREE.LineBasicMaterial({ color: col('--globe-grid'), transparent: true, opacity: num('--globe-grid-a', .12), depthWrite: false }));
  world.add(grat);

  const orbit = new THREE.Group();
  orbit.rotation.set(1.16, 0, .34);
  scene.add(orbit);
  {
    const N = 240, v = [];
    for (let i = 0; i <= N; i++) {
      const a = i / N * Math.PI * 2;
      v.push(Math.cos(a) * R * 1.42, 0, Math.sin(a) * R * 1.42);
    }
    const g = new THREE.BufferGeometry();
    g.setAttribute('position', new THREE.BufferAttribute(new Float32Array(v), 3));
    orbit.add(new THREE.Line(g, new THREE.LineBasicMaterial({
      color: col('--globe-node'), transparent: true, opacity: num('--globe-orbit-a', .3), depthWrite: false })));
  }
  const beads = [0, .38, .71].map((phase, i) => {
    const m = new THREE.Mesh(new THREE.SphereGeometry(.016, 10, 8),
      new THREE.MeshBasicMaterial({ color: col('--globe-node'), transparent: true, opacity: .95 }));
    orbit.add(m);
    return { m, phase, speed: .06 + i * .017 };
  });

  const arcGroup = new THREE.Group(); world.add(arcGroup);
  const heads = [];
  const greatCircle = (a, b, nSteps = 64, arcLift = .28) => {
    const v = ([lon, lat]) => [Math.cos(lat * rad) * Math.cos(lon * rad), Math.sin(lat * rad), Math.cos(lat * rad) * Math.sin(lon * rad)];
    const p = v(a), q = v(b);
    const dotv = Math.max(-1, Math.min(1, p[0] * q[0] + p[1] * q[1] + p[2] * q[2]));
    const om = Math.acos(dotv), s = Math.sin(om) || 1e-6;
    return Array.from({ length: nSteps + 1 }, (_, i) => {
      const t = i / nSteps;
      const k0 = Math.sin((1 - t) * om) / s, k1 = Math.sin(t * om) / s;
      const x = p[0] * k0 + q[0] * k1, y = p[1] * k0 + q[1] * k1, z = p[2] * k0 + q[2] * k1;
      const len = Math.hypot(x, y, z) || 1;
      const lift = 1 + Math.sin(t * Math.PI) * arcLift * Math.sqrt(om / Math.PI);
      return [x / len * lift, y / len * lift, z / len * lift];
    });
  };
  function buildArcs(list) {
    arcGroup.clear(); heads.length = 0;
    list.forEach((r, i) => {
      const a = CITY[r[0] ?? r.from], b = CITY[r[1] ?? r.to];
      if (!a || !b) return;
      const raw = greatCircle(a, b, 72, .58);
      const curve = raw.map(([x, y, z]) => new THREE.Vector3(x, y, -z));
      const c = col(r.level === 'warn' ? '--warn' : '--accent');
      arcGroup.add(new THREE.Line(new THREE.BufferGeometry().setFromPoints(curve),
        new THREE.LineBasicMaterial({ color: c, transparent: true, opacity: .34, depthWrite: false })));
      const hg = new THREE.BufferGeometry();
      hg.setAttribute('position', new THREE.BufferAttribute(new Float32Array(HEAD * 3), 3));
      const head = new THREE.Line(hg, new THREE.LineBasicMaterial({
        color: c, transparent: true, opacity: .95, blending: THREE.AdditiveBlending, depthWrite: false }));
      arcGroup.add(head);
      heads.push({ head, curve, phase: (i * .173) % 1, speed: .16 + (i % 5) * .012 });
    });
  }
  buildArcs(routes);

  const markGroup = new THREE.Group(); world.add(markGroup);
  const sprite = (draw) => {
    const s = 64, c2 = document.createElement('canvas'); c2.width = c2.height = s;
    draw(c2.getContext('2d'), s);
    return new THREE.CanvasTexture(c2);
  };
  const ringTex = sprite((g, s) => {
    g.strokeStyle = '#fff'; g.lineWidth = 3.5;
    g.beginPath(); g.arc(s / 2, s / 2, s / 2 - 3, 0, 7); g.stroke();
  });
  const dotTex = sprite((g, s) => {
    const rg = g.createRadialGradient(s / 2, s / 2, 0, s / 2, s / 2, s / 2);
    rg.addColorStop(0, '#fff'); rg.addColorStop(.3, 'rgba(255,255,255,.9)');
    rg.addColorStop(1, 'rgba(255,255,255,0)');
    g.fillStyle = rg; g.fillRect(0, 0, s, s);
  });
  const pings = [];
  function buildMarkers(list) {
    markGroup.clear(); pings.length = 0;
    list.forEach((m, i) => {
      const p = CITY[m.city]; if (!p) return;
      const v = toVec(p, R * 1.012);
      const c = col(m.level === 'warn' ? '--globe-node-warn' : '--globe-node');
      const dot = new THREE.Sprite(new THREE.SpriteMaterial({
        map: dotTex, color: c, transparent: true, depthWrite: false, blending: THREE.AdditiveBlending }));
      dot.position.copy(v); dot.scale.setScalar(.07 + (m.v || 0) / 100 * .05);
      dot.userData = m;
      markGroup.add(dot);
      const ring = new THREE.Sprite(new THREE.SpriteMaterial({
        map: ringTex, color: c, transparent: true, opacity: .7,
        depthWrite: false, blending: THREE.AdditiveBlending }));
      ring.position.copy(v);
      markGroup.add(ring);
      pings.push({ ring, phase: (i * .29) % 1, base: .06 + (m.v || 0) / 100 * .06 });
    });
  }
  buildMarkers(markers);

  let raf = 0, live = false, t0 = performance.now(), drag = null, vel = 0, seen = true;
  const size = () => {
    const r = host.getBoundingClientRect();
    const w = Math.max(1, r.width), h = Math.max(1, r.height);
    renderer.setSize(w, h, false);
    camera.aspect = w / h; camera.updateProjectionMatrix();
  };

  function tick(now) {
    raf = requestAnimationFrame(tick);
    const t = (now - t0) / 1000;
    if (!drag) { world.rotation.y += (RM.matches ? 0 : spin) * .016 + vel; vel *= .94 }
    const sa = RM.matches ? .6 : t * (Math.PI * 2 / dayLength);
    sun.set(Math.cos(sa), .22, Math.sin(sa)).normalize();

    heads.forEach((h) => {
      const n = h.curve.length, at = (t * h.speed + h.phase) % 1;
      const p = h.head.geometry.attributes.position;
      for (let j = 0; j < HEAD; j++) {
        const u = Math.max(0, at - j * (TAIL / HEAD)) * (n - 1);
        const k = Math.min(n - 2, u | 0), f = u - k;
        const a = h.curve[k], b = h.curve[k + 1];
        p.setXYZ(j, a.x + (b.x - a.x) * f, a.y + (b.y - a.y) * f, a.z + (b.z - a.z) * f);
      }
      p.needsUpdate = true;
      h.head.material.opacity = .95 * Math.min(1, Math.sin(at * Math.PI) * 2.4);
    });

    pings.forEach((p) => {
      const ph = (t * .5 + p.phase) % 1;
      p.ring.scale.setScalar(p.base * (1 + ph * 3.2));
      p.ring.material.opacity = (1 - ph) * .55;
    });

    beads.forEach((b) => {
      const a = (t * b.speed + b.phase) % 1 * Math.PI * 2;
      b.m.position.set(Math.cos(a) * R * 1.42, 0, Math.sin(a) * R * 1.42);
    });

    renderer.render(scene, camera);
  }

  if (interactive) {
    cv.addEventListener('pointerdown', (e) => {
      drag = { x: e.clientX, y: e.clientY }; cv.style.cursor = 'grabbing'; cv.setPointerCapture(e.pointerId);
    });
    cv.addEventListener('pointermove', (e) => {
      if (!drag) return;
      const dx = (e.clientX - drag.x) * .006, dy = (e.clientY - drag.y) * .005;
      world.rotation.y += dx;
      world.rotation.x = Math.max(-1.1, Math.min(1.1, world.rotation.x + dy));
      vel = dx * .35; drag = { x: e.clientX, y: e.clientY };
    });
    const up = () => { drag = null; cv.style.cursor = 'grab' };
    cv.addEventListener('pointerup', up); cv.addEventListener('pointercancel', up);
    if (onPick) {
      const ray = new THREE.Raycaster(); ray.params.Sprite = { threshold: .05 };
      cv.addEventListener('click', (e) => {
        const r = cv.getBoundingClientRect();
        ray.setFromCamera(new THREE.Vector2((e.clientX - r.left) / r.width * 2 - 1,
                                            -((e.clientY - r.top) / r.height) * 2 + 1), camera);
        const hit = ray.intersectObjects(markGroup.children).find((h) => h.object.userData?.city);
        if (hit) onPick(hit.object.userData);
      });
    }
    /* MTC SEAM 4b: hover readout (same raycast path as the click above, on
       pointermove instead) — the MTC dashboard's readout panel shows
       country / USD / share under the cursor; null when nothing is hit. */
    if (onHover) {
      const rayH = new THREE.Raycaster(); rayH.params.Sprite = { threshold: .05 };
      cv.addEventListener('pointermove', (e) => {
        if (drag) return;
        const r = cv.getBoundingClientRect();
        rayH.setFromCamera(new THREE.Vector2((e.clientX - r.left) / r.width * 2 - 1,
                                             -((e.clientY - r.top) / r.height) * 2 + 1), camera);
        const hit = rayH.intersectObjects(markGroup.children).find((h) => h.object.userData?.city);
        cv.style.cursor = hit ? 'pointer' : 'grab';
        onHover(hit ? hit.object.userData : null);
      });
      cv.addEventListener('pointerleave', () => onHover(null));
    }
  }

  const ro = new ResizeObserver(size); ro.observe(host);
  const io = new IntersectionObserver((e) => { seen = e[0].isIntersecting; seen ? api.start() : api.stop() });
  io.observe(host);
  /* A wall screen sits in a background tab for hours: stop cleanly on hide and
     recover if Windows power management drops the SwiftShader context. */
  const vis = () => { document.hidden ? api.stop() : api.start() };
  addEventListener('visibilitychange', vis);
  cv.addEventListener('webglcontextlost', (e) => { e.preventDefault(); api.stop() });
  cv.addEventListener('webglcontextrestored', () => { size(); api.retheme(); api.start() });
  size();

  const api = {
    el: cv, scene, world, camera, renderer,
    start() { if (!live && seen) { live = true; raf = requestAnimationFrame(tick) } },
    stop() { live = false; cancelAnimationFrame(raf) },
    setRoutes(v) { routes = v; buildArcs(v) },
    setMarkers(v) { markers = v; buildMarkers(v) },
    retheme() {
      ocean.material.uniforms.uDeep.value = col('--globe-deep');
      ocean.material.uniforms.uSea.value = col('--globe-sea');
      ocean.material.uniforms.uNight.value = num('--globe-night-k', .32);
      air.material.uniforms.uCol.value = col('--globe-air');
      air.material.uniforms.uA.value = num('--globe-air-a', .9);
      landMat.uniforms.uDay.value = col('--globe-land');
      landMat.uniforms.uDark.value = col('--globe-landnight');
      landMat.uniforms.uCity.value = col('--globe-city');
      coast.material.uniforms.uCol.value = col('--globe-coast');
      coast.material.uniforms.uA.value = num('--globe-coast-a', .85);
      grat.material.color = col('--globe-grid');
      grat.material.opacity = num('--globe-grid-a', .12);
      orbit.children.forEach((o) => { o.material.color = col('--globe-node') });
      buildArcs(routes); buildMarkers(markers);
    },
    focus(city) {
      const p = CITY[city]; if (!p) return;
      const target = -p[0] * rad - Math.PI / 2;
      const cur = world.rotation.y;
      const delta = ((target - cur + Math.PI) % (Math.PI * 2) + Math.PI * 2) % (Math.PI * 2) - Math.PI;
      const t1 = performance.now();
      const ease = () => {
        const k = Math.min(1, (performance.now() - t1) / 900);
        world.rotation.y = cur + delta * (1 - (1 - k) ** 3);
        if (k < 1) requestAnimationFrame(ease);
      };
      RM.matches ? (world.rotation.y = target) : ease();
    },
    destroy() { api.stop(); ro.disconnect(); io.disconnect(); removeEventListener('visibilitychange', vis); renderer.dispose(); cv.remove() },
  };
  addEventListener('kit:theme', () => api.retheme());

  /* Compile shaders off the main thread when the driver can, so the first
     render() doesn't freeze the page. Unsupported drivers fall back silently. */
  const prime = () => { api.start() };
  if (renderer.compileAsync) {
    renderer.compileAsync(scene, camera).then(prime, prime);
  } else prime();
  return api;
}
