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
 *   aria-label      accessible name (default: text of the enclosing figure's .anim-title)
 *
 * Slide data files (slides/slideNN.js) and their images (img/) are loaded from next to this
 * script once a player comes within ~800px of the viewport. Clicking an animated player
 * pauses/resumes it; clicking a static one (no animations) opens it full screen.
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
        // rasters live in img/ next to this script, shared (and cached) across slides
        svgBuf.push('<g data-l="' + k + '">' + L.svg.replace(/@IMG\//g, base + 'img/') + '</g>');
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
        // preload nothing: the poster is shown until the timeline actually starts the video
        v.muted = true; v.loop = true; v.playsInline = true; v.preload = 'none';
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

    // Build the frame right away (slides are 1280x720) so the page never jumps when data arrives.
    var crop = ds.crop ? ds.crop.split(',').map(Number) : [0, 0, 1280, 720];
    this.stage = document.createElement('div');
    this.stage.className = 'ppt-stage';
    this.stage.style.aspectRatio = crop[2] + ' / ' + crop[3];
    root.style.setProperty('--ppt-ar', crop[2] / crop[3]);
    this.inner = document.createElement('div');
    this.inner.className = 'ppt-inner';
    this.inner.style.width = (1280 / crop[2] * 100) + '%';
    this.inner.style.height = (720 / crop[3] * 100) + '%';
    this.inner.style.left = (-crop[0] / crop[2] * 100) + '%';
    this.inner.style.top = (-crop[1] / crop[3] * 100) + '%';
    this.stage.appendChild(this.inner);
    root.appendChild(this.stage);
    if (this.controls) {
      this.bar = document.createElement('div');
      this.bar.className = 'ppt-controls';
      root.appendChild(this.bar);
    }

    // Fetch slide data only when the player comes near the viewport.
    function start() {
      Promise.all(self.ids.map(load)).then(function (datas) { self.init(datas); })
        .catch(function (err) { self.stage.textContent = err.message; });
    }
    if ('IntersectionObserver' in window) {
      var io = new IntersectionObserver(function (entries) {
        if (entries.some(function (e) { return e.isIntersecting; })) { io.disconnect(); start(); }
      }, { rootMargin: '800px 0px' });
      io.observe(root);
    } else {
      start();
    }
  }

  function fullscreenElement() {
    return document.fullscreenElement || document.webkitFullscreenElement;
  }

  Player.prototype.toggleFullscreen = function () {
    var root = this.root;
    if (fullscreenElement() === root) {
      (document.exitFullscreen || document.webkitExitFullscreen).call(document);
    } else {
      var req = root.requestFullscreen || root.webkitRequestFullscreen;
      if (req) req.call(root);
    }
  };

  Player.prototype.init = function (datas) {
    var self = this, root = this.root, stage = this.stage, inner = this.inner;
    root.classList.add('ppt-ready');

    // Players whose slides don't animate (static figures) open full screen on click instead.
    this.animated = datas.some(function (d) { return d.duration > 0; });
    var label = root.getAttribute('aria-label') || (function () {
      var fig = root.closest('figure'), h = fig && fig.querySelector('.anim-title');
      if (!h) return 'Slide';
      h = h.cloneNode(true);
      h.querySelectorAll('.step-no').forEach(function (n) { n.remove(); }); // decorative step badge
      return h.textContent.trim();
    })();
    stage.setAttribute('role', 'img');
    stage.setAttribute('aria-label', label + (this.animated ? ' (animation)' : ''));
    stage.tabIndex = 0;
    if (!this.animated) root.classList.add('ppt-static');

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

    if (this.controls) this.buildControls();
    function activate() { if (self.animated) self.toggle(); else self.toggleFullscreen(); }
    stage.addEventListener('click', activate);
    stage.addEventListener('keydown', function (ev) {
      if (ev.key === ' ' || ev.key === 'Enter') { ev.preventDefault(); activate(); }
    });

    this.render(true);
    if (!this.animated) return; // nothing to play
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

  Player.prototype.buildControls = function () {
    var self = this, bar = this.bar;
    var canFullscreen = document.fullscreenEnabled || document.webkitFullscreenEnabled;
    bar.innerHTML =
      '<button type="button" class="ppt-btn ppt-toggle" aria-label="Play"></button>' +
      '<button type="button" class="ppt-btn ppt-restart" aria-label="Restart">' +
      '<svg viewBox="0 0 24 24"><path d="M12 5V2L7 6l5 4V7a5 5 0 1 1-5 5H5a7 7 0 1 0 7-7z"/></svg></button>' +
      '<div class="ppt-track" role="slider" tabindex="0" aria-label="Seek" aria-valuemin="0" aria-valuemax="100"><div class="ppt-fill"></div>' +
      this.segs.map(function (s) { return '<i style="left:' + (s.start / self.total * 100) + '%"></i>'; }).join('') +
      '</div>' +
      (canFullscreen ? '<button type="button" class="ppt-btn ppt-full" aria-label="Full screen">' +
        '<svg viewBox="0 0 24 24"><path d="M4 9V4h5v2H6v3zm11-5h5v5h-2V6h-3zM4 15h2v3h3v2H4zm14 3v-3h2v5h-5v-2z"/></svg></button>' : '');
    this.btn = bar.querySelector('.ppt-toggle');
    this.fill = bar.querySelector('.ppt-fill');
    this.track = bar.querySelector('.ppt-track');
    var track = this.track;
    this.btn.addEventListener('click', function () { self.toggle(); });
    bar.querySelector('.ppt-restart').addEventListener('click', function () {
      self.seek(0); self.userPaused = false; self.play();
    });
    var full = bar.querySelector('.ppt-full');
    if (full) full.addEventListener('click', function () { self.toggleFullscreen(); });
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
    if (this.fill) {
      var pct = t / this.total * 100;
      this.fill.style.width = pct + '%';
      this.track.setAttribute('aria-valuenow', Math.round(pct));
    }
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
