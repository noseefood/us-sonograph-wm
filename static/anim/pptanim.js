/*
 * pptanim.js — plays PowerPoint slide animations converted to SVG layers + a timeline.
 *
 * Usage:
 *   <link rel="stylesheet" href="static/anim/pptanim.css">
 *   <script src="static/anim/pptanim.js" defer></script>
 *   <div class="ppt-anim" data-slides="slide04,slide05,slide06"></div>
 *
 * Options (data attributes on the .ppt-anim element):
 *   data-slides     comma-separated slide ids, played back to back (required)
 *   data-hold       ms to hold after each slide's animation ends       (default 800)
 *   data-end-hold   ms to hold on the last slide before looping         (default 3000)
 *   data-loop       "false" to stop at the end                          (default true)
 *   data-autoplay   "view" = play while on screen, "none" = wait for a click (default view)
 *   data-controls   "false" to hide the control bar                     (default true)
 *   data-crop       "x,y,w,h" in slide px (1280x720) to show only part of the slide
 *   data-speed      playback rate of the animations, e.g. "1.3"          (default 1;
 *                   holds are given in real ms and are not sped up)
 *
 * Slide data files (slides/slideNN.js) are loaded on demand from next to this script.
 */
(function () {
  'use strict';

  var registry = {};
  var waiting = {};
  var scriptEl = document.currentScript;
  var base = scriptEl ? scriptEl.src.replace(/[^\/]*$/, '') : 'static/anim/';

  function register(id, data) {
    registry[id] = data;
    (waiting[id] || []).forEach(function (cb) { cb(data); });
    delete waiting[id];
  }

  function load(id) {
    return new Promise(function (resolve, reject) {
      if (registry[id]) return resolve(registry[id]);
      if (!waiting[id]) {
        waiting[id] = [];
        var s = document.createElement('script');
        s.src = base + 'slides/' + id + '.js';
        s.onerror = function () { reject(new Error('pptanim: cannot load ' + s.src)); };
        document.head.appendChild(s);
      }
      waiting[id].push(resolve);
    });
  }

  // ---- per-layer state at local time t (pure function of time, so seeking is free) ----
  function pathAt(pts, p) {
    var f = p * (pts.length - 1), i = Math.min(Math.floor(f), pts.length - 2), u = f - i;
    return [pts[i][0] + (pts[i + 1][0] - pts[i][0]) * u, pts[i][1] + (pts[i + 1][1] - pts[i][1]) * u];
  }

  function layerState(L, t) {
    var vis = !L.hidden, op = 1, dx = 0, dy = 0, rot = 0, play = -1;
    var ev = L.ev || [];
    for (var i = 0; i < ev.length; i++) {
      var e = ev[i];
      if (e.t > t) break;
      var p = e.d > 0 ? Math.min(1, (t - e.t) / e.d) : 1;
      switch (e.k) {
        case 'vis': vis = e.v; break;
        case 'fade': vis = true; op = e.v ? p : 1 - p; break;
        case 'move': var xy = pathAt(e.v, p); dx += xy[0]; dy += xy[1]; break;
        case 'rot': rot += e.v * p; break;
        case 'play': play = e.t; break;
      }
    }
    return { vis: vis, op: op, dx: dx, dy: dy, rot: rot, play: play };
  }

  // ---- building the DOM for one slide ----
  function buildSlide(data) {
    var el = document.createElement('div');
    el.className = 'ppt-slide';
    var planes = [], svgBuf = [], layers = [];
    var vb = '0 0 ' + data.w + ' ' + data.h;

    function flushSvg() {
      if (!svgBuf.length) return;
      planes.push({ type: 'svg', html: '<svg class="ppt-plane" viewBox="' + vb + '" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">' + svgBuf.join('') + '</svg>' });
      svgBuf = [];
    }
    data.layers.forEach(function (L, k) {
      if (L.video) {
        flushSvg();
        planes.push({ type: 'video', layer: L, k: k });
      } else {
        svgBuf.push('<g data-l="' + k + '">' + L.svg + '</g>');
      }
    });
    flushSvg();

    planes.forEach(function (pl) {
      if (pl.type === 'svg') {
        el.insertAdjacentHTML('beforeend', pl.html);
      } else {
        var v = document.createElement('video');
        var b = pl.layer.box;
        v.className = 'ppt-video';
        v.src = base + pl.layer.video;
        v.poster = base + pl.layer.poster;
        v.muted = true; v.loop = true; v.playsInline = true; v.preload = 'metadata';
        v.setAttribute('muted', ''); v.setAttribute('playsinline', '');
        v.style.left = (b[0] / data.w * 100) + '%';
        v.style.top = (b[1] / data.h * 100) + '%';
        v.style.width = (b[2] / data.w * 100) + '%';
        v.style.height = (b[3] / data.h * 100) + '%';
        v.dataset.l = pl.k;
        el.appendChild(v);
      }
    });

    el.querySelectorAll('[data-l]').forEach(function (node) {
      var L = data.layers[+node.dataset.l];
      if (!L.ev && !L.hidden && !L.video) return; // static layer: nothing to drive
      layers.push({ L: L, node: node, last: '' });
    });
    return { el: el, data: data, layers: layers };
  }

  function applySlide(S, t, playing, resync) {
    for (var i = 0; i < S.layers.length; i++) {
      var o = S.layers[i], st = layerState(o.L, t), n = o.node;
      if (o.L.video) {
        var want = st.play >= 0 ? (t - st.play) / 1000 : -1;
        if (want < 0) {
          if (!n.paused) n.pause();
          if (n.currentTime !== 0 && resync) n.currentTime = 0;
        } else {
          var dur = n.duration || Infinity;
          if (resync) n.currentTime = want % dur;
          if (playing && n.paused) { var pr = n.play(); if (pr && pr.catch) pr.catch(function () {}); }
          if (!playing && !n.paused) n.pause();
        }
        continue;
      }
      var tr = '';
      if (st.dx || st.dy) tr += 'translate(' + (st.dx * S.data.w).toFixed(2) + ' ' + (st.dy * S.data.h).toFixed(2) + ')';
      if (st.rot) tr += ' rotate(' + st.rot.toFixed(3) + ' ' + o.L.c[0] + ' ' + o.L.c[1] + ')';
      var key = (st.vis ? 1 : 0) + '|' + st.op.toFixed(3) + '|' + tr;
      if (key === o.last) continue;
      o.last = key;
      n.style.display = st.vis ? '' : 'none';
      if (st.op < 1) n.setAttribute('opacity', st.op.toFixed(3)); else n.removeAttribute('opacity');
      if (tr) n.setAttribute('transform', tr); else n.removeAttribute('transform');
    }
  }

  function stopVideos(S) {
    S.el.querySelectorAll('video').forEach(function (v) { if (!v.paused) v.pause(); });
  }

  // ---- player ----
  function Player(root) {
    this.root = root;
    var ds = root.dataset;
    this.ids = (ds.slides || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
    this.hold = ds.hold != null ? +ds.hold : 800;
    this.endHold = ds.endHold != null ? +ds.endHold : 3000;
    this.loop = ds.loop !== 'false';
    this.autoplay = ds.autoplay || 'view';
    this.controls = ds.controls !== 'false';
    this.speed = +ds.speed > 0 ? +ds.speed : 1;
    this.t = 0; this.playing = false; this.visible = false; this.userPaused = false;
    this.reduced = window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches;
    var self = this;
    Promise.all(this.ids.map(load)).then(function (datas) { self.init(datas); })
      .catch(function (err) { root.textContent = err.message; });
  }

  Player.prototype.init = function (datas) {
    var self = this, root = this.root, first = datas[0];
    root.classList.add('ppt-ready');
    var stage = document.createElement('div');
    stage.className = 'ppt-stage';
    var crop = root.dataset.crop ? root.dataset.crop.split(',').map(Number) : null;
    var inner = document.createElement('div');
    inner.className = 'ppt-inner';
    if (crop) {
      stage.style.aspectRatio = crop[2] + ' / ' + crop[3];
      inner.style.width = (first.w / crop[2] * 100) + '%';
      inner.style.height = (first.h / crop[3] * 100) + '%';
      inner.style.left = (-crop[0] / crop[2] * 100) + '%';
      inner.style.top = (-crop[1] / crop[3] * 100) + '%';
    } else {
      stage.style.aspectRatio = first.w + ' / ' + first.h;
    }
    stage.appendChild(inner);
    root.appendChild(stage);

    this.slides = datas.map(function (d) { var S = buildSlide(d); inner.appendChild(S.el); return S; });
    var acc = 0, n = this.slides.length;
    this.segs = this.slides.map(function (S, i) {
      S.el.querySelectorAll('video').forEach(function (v) { v.defaultPlaybackRate = v.playbackRate = self.speed; });
      var len = S.data.duration / self.speed + (i === n - 1 ? self.endHold : self.hold);
      var seg = { S: S, start: acc, len: len };
      acc += len;
      return seg;
    });
    this.total = acc;
    this.active = null;

    if (this.controls) this.buildControls(root);
    stage.addEventListener('click', function () { self.toggle(); });

    this.render(true);
    if (this.reduced) {
      this.seek(this.segs[0].S.data.duration / this.speed);
    } else if (this.autoplay === 'view' && 'IntersectionObserver' in window) {
      new IntersectionObserver(function (entries) {
        entries.forEach(function (e) {
          self.visible = e.isIntersecting;
          if (self.visible && !self.userPaused) self.play(); else if (!self.visible) self.pause(true);
        });
      }, { threshold: 0.35 }).observe(root);
    } else if (this.autoplay !== 'none') {
      this.play();
    }
    document.addEventListener('visibilitychange', function () {
      if (document.hidden) self.pause(true);
      else if (self.visible && !self.userPaused) self.play();
    });
  };

  Player.prototype.buildControls = function (root) {
    var self = this;
    var bar = document.createElement('div');
    bar.className = 'ppt-controls';
    bar.innerHTML =
      '<button type="button" class="ppt-btn ppt-toggle" aria-label="Play"></button>' +
      '<button type="button" class="ppt-btn ppt-restart" aria-label="Restart">' +
      '<svg viewBox="0 0 24 24"><path d="M12 5V2L7 6l5 4V7a5 5 0 1 1-5 5H5a7 7 0 1 0 7-7z"/></svg></button>' +
      '<div class="ppt-track" role="slider" tabindex="0" aria-label="Seek"><div class="ppt-fill"></div>' +
      this.segs.map(function (s) { return '<i style="left:' + (s.start / self.total * 100) + '%"></i>'; }).join('') +
      '</div>';
    root.appendChild(bar);
    this.btn = bar.querySelector('.ppt-toggle');
    this.fill = bar.querySelector('.ppt-fill');
    var track = bar.querySelector('.ppt-track');
    this.btn.addEventListener('click', function () { self.toggle(); });
    bar.querySelector('.ppt-restart').addEventListener('click', function () {
      self.seek(0); self.userPaused = false; self.play();
    });
    function seekFrom(ev) {
      var r = track.getBoundingClientRect();
      var x = Math.min(1, Math.max(0, (ev.clientX - r.left) / r.width));
      self.seek(x * self.total);
    }
    var dragging = false;
    track.addEventListener('pointerdown', function (ev) { dragging = true; track.setPointerCapture(ev.pointerId); seekFrom(ev); });
    track.addEventListener('pointermove', function (ev) { if (dragging) seekFrom(ev); });
    track.addEventListener('pointerup', function () { dragging = false; });
    track.addEventListener('keydown', function (ev) {
      if (ev.key === 'ArrowRight') self.seek(Math.min(self.total, self.t + 1000));
      else if (ev.key === 'ArrowLeft') self.seek(Math.max(0, self.t - 1000));
      else return;
      ev.preventDefault();
    });
    this.updateButton();
  };

  Player.prototype.updateButton = function () {
    if (!this.btn) return;
    this.btn.setAttribute('aria-label', this.playing ? 'Pause' : 'Play');
    this.btn.innerHTML = this.playing
      ? '<svg viewBox="0 0 24 24"><path d="M7 5h4v14H7zM13 5h4v14h-4z"/></svg>'
      : '<svg viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>';
  };

  Player.prototype.toggle = function () {
    if (this.playing) { this.userPaused = true; this.pause(); }
    else {
      this.userPaused = false;
      if (this.t >= this.total) this.seek(0);
      this.play();
    }
  };

  Player.prototype.play = function () {
    if (this.playing || !this.segs) return;
    this.playing = true;
    this.updateButton();
    var self = this, prev = performance.now();
    this.render(false);
    function tick(now) {
      if (!self.playing) return;
      self.t += Math.min(100, now - prev);
      prev = now;
      if (self.t >= self.total) {
        if (self.loop) { self.t = 0; self.render(true); }
        else { self.t = self.total; self.render(false); self.pause(); return; }
      }
      self.render(false);
      self.raf = requestAnimationFrame(tick);
    }
    this.raf = requestAnimationFrame(tick);
  };

  Player.prototype.pause = function () {
    if (!this.segs) return;
    this.playing = false;
    cancelAnimationFrame(this.raf);
    if (this.active) stopVideos(this.active.S);
    this.updateButton();
  };

  Player.prototype.seek = function (t) {
    this.t = t;
    this.render(true);
  };

  Player.prototype.render = function (resync) {
    var t = this.t, seg = this.segs[this.segs.length - 1];
    for (var i = 0; i < this.segs.length; i++) {
      if (t < this.segs[i].start + this.segs[i].len) { seg = this.segs[i]; break; }
    }
    if (seg !== this.active) {
      if (this.active) { this.active.S.el.classList.remove('ppt-on'); stopVideos(this.active.S); }
      seg.S.el.classList.add('ppt-on');
      this.active = seg;
      resync = true;
    }
    // player time is real ms; slide time runs `speed` times faster
    applySlide(seg.S, Math.min((t - seg.start) * this.speed, seg.S.data.duration + 1), this.playing, resync);
    if (this.fill) this.fill.style.width = (t / this.total * 100) + '%';
  };

  function mountAll() {
    document.querySelectorAll('.ppt-anim:not([data-ppt-mounted])').forEach(function (el) {
      el.setAttribute('data-ppt-mounted', '');
      el.pptPlayer = new Player(el);
    });
  }

  window.PPTAnim = { register: register, mount: mountAll, load: load };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', mountAll);
  else mountAll();
})();
