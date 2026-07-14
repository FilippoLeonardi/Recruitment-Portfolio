/* ════════════════════════════════════════════════
   ENGINEERING PORTFOLIO — MAIN.JS
   ════════════════════════════════════════════════ */

'use strict';

/* ══════════ 1. LOADER ══════════ */
window.addEventListener('load', () => {
  setTimeout(() => {
    const l = document.getElementById('loader');
    if (l) l.classList.add('hide');
  }, 600);
});

/* ══════════ 2. PLACEHOLDER ICONS ══════════
   Font Awesome uses pseudo-elements so we inject
   a real <i> inside each .ph div instead.         */
document.querySelectorAll('.ph').forEach(el => {
  const icon  = el.dataset.icon  || 'fa-image';
  const label = el.dataset.label || '';
  el.innerHTML = `
    <i class="fas ${icon}" style="font-size:2.8rem;color:rgba(255,255,255,.2);position:relative;z-index:1"></i>
    ${label ? `<span style="font-size:.72rem;color:rgba(255,255,255,.18);letter-spacing:1px;position:relative;z-index:1">${label}</span>` : ''}
  `;
});

/* ══════════ 3. PARTICLES ══════════ */
const Particles = (() => {
  let canvas, ctx, list = [], mouse = { x: null, y: null }, raf;
  const CFG = { n: 55, maxD: 130, spd: 0.35, r: 1.8, connAlpha: 0.14 };

  function init() {
    canvas = document.getElementById('particlesCanvas');
    if (!canvas) return;
    ctx = canvas.getContext('2d');
    resize(); build();
    window.addEventListener('resize', () => { resize(); build(); });
    window.addEventListener('mousemove', e => { mouse.x = e.clientX; mouse.y = e.clientY; });
    window.addEventListener('mouseleave', () => { mouse.x = null; mouse.y = null; });
    loop();
  }

  function resize() {
    canvas.width  = window.innerWidth;
    canvas.height = window.innerHeight;
  }

  function build() {
    list = Array.from({ length: CFG.n }, () => ({
      x:  Math.random() * canvas.width,
      y:  Math.random() * canvas.height,
      vx: (Math.random() - .5) * CFG.spd,
      vy: (Math.random() - .5) * CFG.spd,
      r:  Math.random() * CFG.r + .8,
      a:  Math.random() * .4 + .2,
    }));
  }

  function draw() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    /* connections */
    for (let i = 0; i < list.length; i++) {
      for (let j = i + 1; j < list.length; j++) {
        const dx = list[i].x - list[j].x, dy = list[i].y - list[j].y;
        const d  = Math.hypot(dx, dy);
        if (d < CFG.maxD) {
          ctx.beginPath();
          ctx.moveTo(list[i].x, list[i].y);
          ctx.lineTo(list[j].x, list[j].y);
          ctx.strokeStyle = `rgba(79,142,247,${(1 - d / CFG.maxD) * CFG.connAlpha})`;
          ctx.lineWidth = .8;
          ctx.stroke();
        }
      }
      /* mouse connection */
      if (mouse.x) {
        const dm = Math.hypot(list[i].x - mouse.x, list[i].y - mouse.y);
        if (dm < 160) {
          ctx.beginPath();
          ctx.moveTo(list[i].x, list[i].y);
          ctx.lineTo(mouse.x, mouse.y);
          ctx.strokeStyle = `rgba(34,211,238,${(1 - dm / 160) * .38})`;
          ctx.lineWidth = .9;
          ctx.stroke();
        }
      }
    }

    /* dots */
    list.forEach(p => {
      ctx.beginPath();
      ctx.arc(p.x, p.y, p.r, 0, Math.PI * 2);
      ctx.fillStyle = `rgba(79,142,247,${p.a})`;
      ctx.fill();
    });
  }

  function update() {
    list.forEach(p => {
      p.x += p.vx; p.y += p.vy;
      if (p.x < 0 || p.x > canvas.width)  p.vx *= -1;
      if (p.y < 0 || p.y > canvas.height) p.vy *= -1;
      /* mouse repulsion */
      if (mouse.x) {
        const dx = p.x - mouse.x, dy = p.y - mouse.y;
        const d  = Math.hypot(dx, dy);
        if (d < 75) {
          const f = (75 - d) / 75 * .28;
          p.vx += (dx / d) * f;
          p.vy += (dy / d) * f;
          const spd = Math.hypot(p.vx, p.vy);
          if (spd > 2) { p.vx = p.vx / spd * 2; p.vy = p.vy / spd * 2; }
        }
      }
    });
  }

  function loop() { draw(); update(); raf = requestAnimationFrame(loop); }

  return { init };
})();

/* ══════════ 5. SCROLL ANIMATIONS ══════════ */
const AOS = (() => {
  function init() {
    const els = document.querySelectorAll('[data-aos]');
    const obs = new IntersectionObserver((entries) => {
      entries.forEach(e => {
        if (e.isIntersecting) {
          const d = +e.target.dataset.aosDelay || 0;
          setTimeout(() => e.target.classList.add('in'), d);
        }
      });
    }, { threshold: .1, rootMargin: '0px 0px -40px 0px' });
    els.forEach(el => obs.observe(el));
  }
  return { init };
})();

/* ══════════ 6. COUNTER ══════════ */
const Counter = (() => {
  function run(el) {
    const target = +el.dataset.target, dur = 1800, step = target / (dur / 16);
    let cur = 0;
    const t = setInterval(() => {
      cur += step;
      if (cur >= target) { cur = target; clearInterval(t); }
      el.textContent = Math.floor(cur);
    }, 16);
  }
  function init() {
    const els = document.querySelectorAll('.cnt');
    const obs = new IntersectionObserver(entries => {
      entries.forEach(e => { if (e.isIntersecting && !e.target._counted) { e.target._counted = true; run(e.target); } });
    }, { threshold: .5 });
    els.forEach(el => obs.observe(el));
  }
  return { init };
})();

/* ══════════ 7. NAVBAR ══════════ */
const Nav = (() => {
  let nav, ham, menu, links, sections;

  function init() {
    nav      = document.getElementById('navbar');
    ham      = document.getElementById('hamburger');
    menu     = document.getElementById('navLinks');
    links    = document.querySelectorAll('.nav-link');
    sections = document.querySelectorAll('section[id]');

    window.addEventListener('scroll', onScroll, { passive: true });
    ham?.addEventListener('click', () => {
      ham.classList.toggle('open');
      menu.classList.toggle('open');
    });
    links.forEach(a => a.addEventListener('click', () => {
      ham?.classList.remove('open');
      menu?.classList.remove('open');
    }));
    onScroll();
  }

  function onScroll() {
    /* Sticky bg */
    nav.classList.toggle('scrolled', window.scrollY > 20);
    /* Back-to-top */
    document.getElementById('backToTop')?.classList.toggle('show', window.scrollY > 500);
    /* Active link */
    let cur = '';
    sections.forEach(s => {
      if (window.scrollY >= s.offsetTop - 90) cur = s.id;
    });
    links.forEach(a => {
      a.classList.toggle('active', a.getAttribute('href') === `#${cur}`);
    });
  }

  return { init };
})();

/* ══════════ 8. THEME TOGGLE ══════════ */
const Theme = (() => {
  function apply(t) {
    document.documentElement.dataset.theme = t;
    localStorage.setItem('pf-theme', t);
    const ic = document.getElementById('themeIcon');
    if (ic) ic.className = t === 'dark' ? 'fas fa-sun' : 'fas fa-moon';
  }
  function init() {
    apply(localStorage.getItem('pf-theme') || 'dark');
    document.getElementById('themeToggle')?.addEventListener('click', () => {
      apply(document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark');
    });
  }
  return { init };
})();

/* ══════════ 9. TABS ══════════ */
const Tabs = (() => {
  function init() {
    document.querySelectorAll('.tab-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
        btn.classList.add('active');
        const panel = document.getElementById(`${btn.dataset.tab}-tab`);
        if (panel) {
          panel.classList.add('active');
          panel.querySelectorAll('[data-aos]').forEach(el => el.classList.add('in'));
        }
      });
    });
  }
  return { init };
})();

/* ══════════ 10. CARD TILT ══════════ */
const Tilt = (() => {
  function init() {
    document.querySelectorAll('.proj-card, .res-card').forEach(card => {
      card.addEventListener('mousemove', e => {
        const r  = card.getBoundingClientRect();
        const rx = ((e.clientY - r.top)  / r.height - .5) * -10;
        const ry = ((e.clientX - r.left) / r.width  - .5) *  10;
        card.style.transform = `perspective(900px) rotateX(${rx}deg) rotateY(${ry}deg) translateY(-9px)`;
      });
      card.addEventListener('mouseleave', () => { card.style.transform = ''; });
    });
  }
  return { init };
})();

/* ══════════ 11. MODALS ══════════ */
const MODAL_DATA = {
  p1: {
    hue:'220', year:'2026', image:'static/images/LBO_background.jpeg',
    title:'Automated LBO Valuation Model',
    tags:[['Python','t-blue'],['LBO Modeling','t-cyan'],['Financial Modeling','t-purple'],['Excel','t-green']],
    desc:'Designed and implemented a Python-powered LBO tool that requires only essential financial assumptions to automatically generate a fully populated Excel model, automating valuation, debt schedules, cash flow projections, and investor return calculations. ',
    highlights:[
      'Automated full LBO model generation from fewer than 15 essential financial and transaction assumptions',
      'Modeled revenue growth, EBITDA margins, working capital, CapEx, D&A, and debt repayment schedules',
      'Calculated purchase price, enterprise value, equity contribution, cash flow available for debt paydown, MOIC, and IRR',
      'Streamlined private equity valuation workflow by eliminating manual entry across 100+ linked Excel cells',
    ],
    calc:'/Projects/lbo_calculator/index.html',
  },
  p2: {
    hue:'265', year:'2026', image:'static/images/Certificate.jpeg',
    title:'Financial Modeling and Valuation Intensive',
    tags:[['Valuation Analysis','t-green'],['Three-Statement Forecasting','t-orange'],['Corporate Finance','t-blue']],
    desc:'Completed an intensive financial modeling course with City Investment Training, covering accounting, three-statement modeling, valuation, LBO analysis, and M&A accretion/dilution, including a full Burger King case study model built from first principles.',
    highlights:[
      'Built fully integrated 3-statement financial model',
      'Constructed PP&E, working capital, debt, and shareholders’ equity schedules',
      'Performed DCF, LBO, Trading Comps, and Transaction Comps valuation analyses',
      'Completed M&A accretion/dilution analysis to assess transaction impact',
    ],
    calc:'/Projects/xlsx_viewer/index.html',
    calcLabel:'Visualize work',
    cert:'static/images/Excel%20Certification%20copy.pdf',
    certLabel:'Excel Certificate',
  },

  p3: {
    hue:'265', year:'2026', image:'static/images/NN+Poly.jpeg',
    title:'ML Optimization Tool',
    tags:[['Python','t-green'],['Neural Network Training','t-orange'],['Mathematical Modeling and Statistics','t-blue']],
    desc:'Developed a hybrid machine learning model in Python to predict aerodynamic coefficients from vehicle state parameters, replacing computationally expensive Kriging interpolation with a polynomial regression corrected by a neural network architecture that maintained <1% prediction error comapred to the validated Kriging model.',
    highlights:[
      'Obtained a 53% run-time reduction with my model comapred to the standard Kriging interpolation method',
      'Acheived <1% prediction error with respect to the Kriging interpolation baseline when predicting aerodynamic coefficients for unseen vehicle states',
      'Handles 200+ aerodynamic variables concurrently.',
      'Implemented into the developing simulation framework of the department, with the goal of predicting wind tunnel results, eventually replacing the need for physical wind tunnel testing and significantly reducing costs and time-to-market.',
    ],
    private:true,
  },
  
  p4: {
    hue:'185', year:'2026', image:'static/images/robot ES 51 /IMG_4072.jpg',
    title:'ES-51 Embedded Systems Robotics Project', 
    tags:[['CAD','t-red'],['Electronics','t-blue'],['Manufacturing','t-green'],['Prototyping','t-cyan']],
    desc:'End-to-end mechatronic robot built from scratch over a semester for Harvard\'s ES-51 course. Designed a custom gearbox and drivetrain around a DC motor, engineered a full chassis in CAD, fabricated every part in-house, and validated the design with quantitative mechanical analysis.',
    highlights:[
      'Custom gearbox design — selected gear ratios to balance torque and top speed against motor specs, sizing the reduction for reliable climbing and traction on the ramps.',
      'Full CAD assembly modeled in SolidWorks — chassis, drivetrain, and mechanism designed as a mated assembly with proper tolerances, clearances, and fits.',
      'Real-time multi-sensor integration — implemented a microcontroller-based control system to read from multiple sensors and actuate the drivetrain in real time.',
      '2nd place — Turf Wars competition (15 teams)',
    ],
    images:[
      'static/images/robot ES 51 /IMG_4242 5.jpeg',
      'static/images/robot ES 51 /IMG_4243 3.jpeg',
      'static/images/robot ES 51 /IMG_4283 4.jpeg',
      'static/images/robot ES 51 /IMG_4299 4.jpeg',
      'static/images/robot ES 51 /IMG_4300 4.jpeg',
      'static/images/robot ES 51 /IMG_4301 4.jpeg',
      'static/images/robot ES 51 /IMG_4361 4.jpeg',
      'static/images/robot ES 51 /IMG_4370 4.jpeg',
      'static/images/robot ES 51 /IMG_4371 4.jpeg',
      'static/images/robot ES 51 /IMG_4373 4.jpeg',
      'static/images/robot ES 51 /IMG_4440 5.jpeg',
      'static/images/robot ES 51 /IMG_4465 4.jpeg',
      'static/images/robot ES 51 /IMG_4478 5.jpeg',
      'static/images/robot ES 51 /Robot img 1.jpeg',
      'static/images/robot ES 51 /Robot img 2.jpeg',
      'static/images/robot ES 51 /IMG_5988 4.MP4',
    ],
    calc:'/Projects/cad_viewer/index.html',
  },


  p5: {
    hue:'265', year:'2026', image:'static/images/PlotOne.jpeg',
    title:'PlotOne - Interactive Aerodynamics Data Visualization Platform',
    tags:[['Python','t-green'],['Data Analysis and Interpretation','t-orange'],['Signal Processing','t-blue'],['Numerical Methods','t-red'],['GUI development','t-yellow'],['Optimization Workflows','t-purple'], ['Analytical Thinking','t-cyan'], ['Problem Solving','t-pink'], ['Collaboration','t-green']],
    desc:'Developed a Python data visualization platform for vehicle dynamics engineers to analyze large aerodynamic and simulation datasets without relying on external software such as PlotJuggler or MoTeC. It supports interactive 2D and 3D plotting directly from CSV files, advanced mathematical manipulation of signals, session persistence, and customizable plotting environments, allowing engineers to inspect simulation outputs and iterate on vehicle models significantly faster.',
    highlights:[
      'Supports 2D/3D plotting, advanced signal filtering and gating, collaborative external mathematical function libraries, and real-time polynomial interpolation of scatter data',
      'Analysis of multiple telemetry and simulation datasets (CSV files) simultaneously, with the ability to overlay and compare results from different simulation runs or experimental tests',
      'Reduced engineering data analysis and simulation validation time by approximately 40% by integrating all visualization directly into the optimization workflow',
      'Adopted into the Vehicle Dynamics Department\'s primary optimization framework after only 3 days of pilot testing, providing engineers with instant graphical feedback after every optimization run',
      'Engineered a scalable multi-window visualization architecture supporting up to 25 synchronized plots and four concurrent signal overlays per graph'
    ],
    private:true,
  },


  p6: {
    hue:'265', year:'2026', image:'static/images/Form SAE/Logo_2.png',
    title:'Formula SAE - Crimson Motorsports',
    tags:[['Suspension Design','t-blue'],['CAD & FEA','t-green'],['Vehicle Dynamics','t-orange'],['Technical Leadership','t-cyan'],['Project Management','t-purple'],['Supplier Negotiation','t-red'], ['Strategic Decision-Making','t-yellow'], ['Cross-Functional Collaboration','t-pink']],
    desc:'As Mechanical Engineering Team Lead for Harvard\'s Formula SAE team, I led the design and development of the vehicle\'s suspension system, working alongside a senior thesis students to develop and validate key suspension components. My work combined analytical modeling, CAD design, and engineering trade-off analysis, along with budget management, supplier negotiations, and procurement planning to deliver a cost-effective design while maximizing engineering performance.',
    highlights:[
      'Actively managed a $20,000 manufacturing and procurement budget, negotiated supplier discounts, secured corporate sponsorships, and optimized component sourcing to reduce manufacturing costs by approximately 15% while keeping the project on schedule',
      'Led a multidisciplinary team of 10+ engineers responsible for suspension development, coordinating design reviews, task allocation, and integration with the chassis, steering, and powertrain teams throughout the vehicle development cycle.',
      'Designed and validated a double-wishbone suspension geometry, developing bespoke mounting solutions and suspension members.',
      'Collaborated with a senior thesis student to engineer the vehicle\'s rocker and pushrod suspension system, evaluating motion ratios, wheel-rate characteristics and manufacturability.',
    ],
    images:[
      'static/images/Form SAE/IMG_4776.jpg',
      'static/images/Form SAE/IMG_4778.jpg',
      'static/images/Form SAE/IMG_4780.jpg',
      'static/images/Form SAE/IMG_4535.jpeg',
      'static/images/Form SAE/IMG_4536.jpeg',
      'static/images/Form SAE/IMG_4537.jpeg',
      'static/images/Form SAE/IMG_4538.jpeg',
      'static/images/Form SAE/IMG_4472.jpeg',
      'static/images/Form SAE/40756807-1517-4b86-889f-7b3c44b9d7b9.JPG',
      'static/images/Form SAE/IMG_4428.jpeg',
      'static/images/Form SAE/IMG_4429.jpeg',
      'static/images/Form SAE/IMG_4430.jpeg',
      'static/images/Form SAE/IMG_4431.jpeg',
      'static/images/Form SAE/IMG_4772.jpg',
      'static/images/Form SAE/IMG_4773.jpg',
      'static/images/Form SAE/IMG_4775.jpg',
    ],
    imagesLabel:'Build Gallery',
  },


  p7: {
    hue:'265', year:'2025', image:'static/images/Robotics images/mars-rover.jpg',
    imageFramed:true,
    title:'Harvard Undergraduate Robotics Club',
    tags:[['Robotic Manipulation','t-blue'],['Mechanism Design','t-red'],['Motion Planning','t-green'],['Design Optimization','t-orange'],['Systems Integration','t-purple'],['Interdisciplinary Collaboration','t-cyan'],['Systems Thinking','t-yellow'],['Fundraising','t-green']],
    desc:'As a Mechanical Systems Developer for Harvard\'s Mars Rover Team, I contributed to the mechanical design, analysis, and validation of the rover\'s mobility, suspension and manipulation systems. I focused on optimizing the suspension architecture for better performance and stability on rough surfaces, while also contributing to the design of tge 6-degree-of-freedom robotic arm. The project involved iterative CAD design, finite element analysis (FEA), rapid prototyping, and subsystem integration with the electrical and software teams, alongside managing sponsor outreach and fundraising initiatives to support the team\'s manufacturing and competition budget.',
    highlights:[
      'Redesigned the rover\'s rocker-bogie suspension system, changing from 6- to 4-wheel configuration using kinematic and load-distribution analyses, improving mars-like surface navigation while also reducing overall system mass by ~12%.',
      'Designed and tested a 6-degree-of-freedom robotic arm, optimizing joint geometry and structural stiffness while integrating actuators to maximize payload capacity and precision.',
      'Developed and cadded housing structures for the rover\'s electronics and battery systems, focusing on easy accessibility for maintainence and upgrades',
      'Contributed to sponsorship and fundraising initiatives by preparing technical proposals, engaging with corporate partners and private donors, and helping secure over $10,000 from various sources.',
    ],
    images:[
      'static/images/Robotics images/Video 1.mp4',
      'static/images/Robotics images/image 2.jpg',
      'static/images/Robotics images/image 3.jpg',
      'static/images/Robotics images/Video 4.mp4',
      'static/images/Robotics images/Image 5.jpg',
      'static/images/Robotics images/Image 6.jpg',
      'static/images/Robotics images/Image 7.jpg',
      'static/images/Robotics images/Image 8.jpg',
      'static/images/Robotics images/Image 9.jpg',
      'static/images/Robotics images/Image 10.jpg',
      'static/images/Robotics images/Image 11.jpg',
      'static/images/Robotics images/Image 12.jpg',
      'static/images/Robotics images/Image 13.jpg',
      'static/images/Robotics images/Video 14.mp4',
    ],
    imagesLabel:'See the Rover in Action',
  },


  p8: {
    hue:'265', year:'2025', image:'static/images/Machine Shop/Home_pic.jpeg',
    imageBorder:true,
    title:'Lucy Engine Manufacturing and Assembly',
    tags:[['Precision Machining','t-blue'],['CNC Manufacturing','t-red'],['Manual Lathe Operation','t-green'],['Metrology & Tolerancing','t-orange'],['Technical Drawing Interpretation','t-purple'],['Attention to Detail','t-cyan'],['Problem Solving','t-yellow'],['Manufacturing Planning','t-red']],
    desc:'Built a fully functional Lucy reciprocating engine from raw stock materials, independently manufacturing every component from aluminum, brass, steel, and acrylic using manual and CNC machine tools. Working from detailed engineering drawings, I applied precision machining techniques—including turning, facing, milling, boring, drilling, reaming, tapping, threading, and surface finishing—to produce tight-tolerance components before performing complete mechanical assembly and functional validation of the engine. The project provided hands-on experience in manufacturing processes, geometric tolerancing, metrology, and design-for-manufacture principles while developing a deep understanding of mechanical systems and precision engineering.',
    highlights:[
      'Manufactured and assembled 11 precision-engineered components entirely from raw aluminum, brass, steel, and acrylic stock, achieving dimensional tolerances as tight as ±0.001 in on critical features.',
      'Operated 10+ industrial manufacturing machines, including manual and CNC lathes, Bridgeport and TRAK milling machines, drill presses, surface grinders, and band saws to perform turning, facing, boring, drilling, reaming, tapping and threading.',
      'Interpreted and manufactured from 14 detailed engineering drawings, translating GD&T requirements, machining sequences, and assembly specifications into finished components while maintaining dimensional accuracy',
      'Optimized material utilization and machining strategy, reducing material waste by ~20% and minimizing manufacturing costs while maintaining tight dimensional tolerances',
    ],
    images:[
      'static/images/Machine Shop/Home_pic.jpeg',
      'static/images/Machine Shop/IMG_4551.jpeg',
      'static/images/Machine Shop/IMG_4552.jpeg',
      'static/images/Machine Shop/IMG_4553.jpeg',
      'static/images/Machine Shop/IMG_4554.jpeg',
      'static/images/Machine Shop/IMG_4555.jpeg',
      'static/images/Machine Shop/IMG_4556.jpeg',
      'static/images/Machine Shop/IMG_4557.jpeg',
      'static/images/Machine Shop/IMG_4566.jpeg',
    ],
    imagesLabel:'Final Result',
  },

  r1: {
    hue:'220', year:'2025', image:'static/images/RC_circuit.jpg',
    title:'Prediction and Analysis of Low-Pass and High-Pass RC Filters in Musical Note Isolation',
    tags:[['Physics','t-blue'],['Signal Processing','t-cyan'],['IB Extended Essay','t-purple']],
    desc:'Research Question: How can Resistance and Capacitance be used to Create Low-Pass and High-Pass Filters to Detect the Notes Making up a Chord?<br><br>This research paper investigates the relationship between cutoff frequency and gain in RC low-pass and high-pass filters, and how these principles can be used to isolate the individual notes within a musical chord. By combining circuit theory with experimental signal processing, recorded piano chords were converted into electrical signals and analyzed using filters with varying resistance and capacitance values. The results demonstrate how filter frequency response selectively attenuates or passes specific frequency components, enabling complex audio signals to be decomposed into their constituent notes. The project highlights the practical application of analog filters in audio engineering and signal processing while providing an intuitive demonstration of fundamental concepts in wave physics and electronic circuits.',
    paper:'static/images/FINAL%20-%20Leonardi%2C%20Filippo%20-%20lmb889%20-%20Physics-%20EE%22%20copy.pdf',
  },
};

function openModal(id) {
  const d = MODAL_DATA[id];
  if (!d) return;

  const tags = d.tags.map(([t, c]) => `<span class="tag ${c}">${t}</span>`).join('');
  const hls  = (d.highlights || []).map(h => `
    <div class="hl-item"><i class="fas fa-check-circle"></i><span>${h}</span></div>`).join('');
  const hlSection = (d.highlights && d.highlights.length) ? `
      <div class="modal-hl">
        <h4><i class="fas fa-star" style="color:var(--orange);margin-right:6px"></i>Key Highlights</h4>
        <div class="hl-list">${hls}</div>
      </div>` : '';

  const acts = [];
  if (d.images) {
    acts.push(`<a href="#" class="btn btn-primary" onclick="openGallery(MODAL_DATA['${id}'].images); return false;"><i class="fas fa-images"></i> ${d.imagesLabel || 'Images'}</a>`);
  } else if (d.gh) {
    acts.push(`<a href="${d.gh}" class="btn btn-primary" target="_blank" rel="noopener"><i class="fab fa-github"></i> GitHub</a>`);
  }
  if (d.calc) {
    acts.push(`<a href="#" class="btn btn-outline" onclick="closeModal(); openCalculator('${d.calc}'); return false;"><i class="fas fa-calculator"></i> ${d.calcLabel || 'Live Demo'}</a>`);
  } else if (d.private) {
    acts.push(`<a href="#projects" class="btn btn-outline" title="Built with private company data — can't be shared outside the firm." onclick="closeModal();"><i class="fas fa-lock"></i> Built with private company data — unfortunately can't be shared outside the firm</a>`);
  } else if (d.demo) {
    acts.push(`<a href="${d.demo}" class="btn btn-outline" target="_blank" rel="noopener"><i class="fas fa-external-link-alt"></i> Live Demo</a>`);
  }
  if (d.paper) {
    const paperTitle = (d.paperTitle || d.title || '').replace(/'/g, "\\'");
    acts.push(`<a href="#" class="btn btn-outline" onclick="closeModal(); openDocViewer('${d.paper}', '${paperTitle}'); return false;"><i class="fas fa-file-pdf"></i> Read Paper</a>`);
  }
  if (d.cert) {
    const certTitle = (d.certLabel || 'Certificate').replace(/'/g, "\\'");
    acts.push(`<a href="#" class="btn btn-outline" onclick="closeModal(); openDocViewer('${d.cert}', '${certTitle}'); return false;"><i class="fas fa-file-pdf"></i> ${d.certLabel || 'Certificate'}</a>`);
  }

  document.getElementById('modalContent').innerHTML = `
    <img src="${d.image}" alt="${d.title}" class="modal-image${d.imageContain ? ' contain' : ''}${d.imagePadLeft ? ' pad-left' : ''}${d.imageBorder ? ' bordered' : ''}${d.imageFramed ? ' framed' : ''}">

    <div class="modal-body">
      <div class="tag-row">${tags}<span class="tag t-blue">${d.year}</span></div>
      <h2 class="modal-title">${d.title}</h2>
      <p class="modal-text">${d.desc}</p>
${hlSection}
      <div class="modal-acts">
        ${acts.join('')}
      </div>
    </div>
  `;

  document.getElementById('modalOverlay').classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeModal() {
  document.getElementById('modalOverlay').classList.remove('open');
  document.body.style.overflow = '';
}

document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

let galleryImages = [];
let galleryIndex = 0;

function isVideoPath(path) {
  return /\.(mp4|mov|webm|m4v)$/i.test(path);
}

function renderGalleryImage() {
  const path = galleryImages[galleryIndex];
  const $img = document.getElementById('galleryImg');
  const $video = document.getElementById('galleryVideo');

  $video.pause();
  if (isVideoPath(path)) {
    $video.src = path;
    $video.style.display = 'block';
    $img.style.display = 'none';
    $img.removeAttribute('src');
  } else {
    $video.removeAttribute('src');
    $video.load();
    $video.style.display = 'none';
    $img.style.display = 'block';
    $img.src = path;
  }
  document.getElementById('galleryCounter').textContent = (galleryIndex + 1) + ' / ' + galleryImages.length;
}

function openGallery(images) {
  galleryImages = images;
  galleryIndex = 0;
  renderGalleryImage();
  document.getElementById('galleryOverlay').classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeGallery() {
  document.getElementById('galleryOverlay').classList.remove('open');
  document.getElementById('galleryVideo').pause();
  document.body.style.overflow = '';
}

function galleryStep(delta) {
  galleryIndex = (galleryIndex + delta + galleryImages.length) % galleryImages.length;
  renderGalleryImage();
}

document.addEventListener('keydown', e => {
  if (!document.getElementById('galleryOverlay').classList.contains('open')) return;
  if (e.key === 'Escape') closeGallery();
  if (e.key === 'ArrowRight') galleryStep(1);
  if (e.key === 'ArrowLeft') galleryStep(-1);
});

function openDocViewer(url, title) {
  // Route through our own PDF.js-based viewer instead of pointing the
  // iframe straight at the PDF — some browsers are configured to always
  // download PDFs instead of showing them inline, which would skip the
  // preview entirely. Rendering it ourselves guarantees it always opens
  // in-page first; the Download button below still links to the real file.
  const absoluteUrl = new URL(url, document.baseURI).href;
  document.getElementById('docFrame').src = '/Projects/pdf_viewer/index.html?file=' + encodeURIComponent(absoluteUrl);
  document.getElementById('docTitle').textContent = title || '';
  const dl = document.getElementById('docDownload');
  dl.href = url;
  dl.setAttribute('download', (title || 'document') + '.pdf');
  document.getElementById('docOverlay').classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeDocViewer() {
  document.getElementById('docOverlay').classList.remove('open');
  document.getElementById('docFrame').src = '';
  document.body.style.overflow = '';
}

document.addEventListener('keydown', e => { if (e.key === 'Escape') closeDocViewer(); });




function openCalculator(url) {
  const frame = document.getElementById('calcFrame');
  frame.src = url;
  document.getElementById('calcOverlay').classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeCalculator() {
  document.getElementById('calcOverlay').classList.remove('open');
  document.getElementById('calcFrame').src = '';
  document.body.style.overflow = '';
}





document.addEventListener('keydown', e => { if (e.key === 'Escape') closeCalculator(); });

/* ══════════ 12. CONTACT FORM ══════════ */
const Form = (() => {
  function init() {
    document.getElementById('contactForm')?.addEventListener('submit', async e => {
      e.preventDefault();
      const btn     = document.getElementById('submitBtn');
      const txt     = btn.querySelector('.btn-txt');
      const load    = btn.querySelector('.btn-load');
      const msg     = document.getElementById('formMsg');

      txt.classList.add('hidden');
      load.classList.remove('hidden');
      btn.disabled = true;
      msg.className = 'form-msg';

      const body = {
        name:    document.getElementById('fname').value.trim(),
        email:   document.getElementById('femail').value.trim(),
        subject: document.getElementById('fsubject').value.trim(),
        message: document.getElementById('fmessage').value.trim(),
      };

      try {
        const res = await fetch('/api/contact', {
          method:'POST',
          headers:{ 'Content-Type':'application/json' },
          body:JSON.stringify(body),
        });
        const data = await res.json();
        msg.textContent = data.message;
        msg.classList.add(data.status);
        if (data.status === 'success') {
          document.getElementById('contactForm').reset();
          confetti();
        }
      } catch {
        msg.textContent = 'Neterror — please try again.';
        msg.classList.add('error');
      } finally {
        txt.classList.remove('hidden');
        load.classList.add('hidden');
        btn.disabled = false;
      }
    });
  }

  /* tiny confetti burst */
  function confetti() {
    const colors = ['#4f8ef7','#22d3ee','#34d399','#fbbf24','#a78bfa'];
    for (let i = 0; i < 24; i++) {
      const dot = document.createElement('div');
      const sz  = Math.random() * 7 + 4;
      dot.style.cssText = `
        position:fixed;width:${sz}px;height:${sz}px;
        background:${colors[i % colors.length]};border-radius:50%;
        pointer-events:none;z-index:99999;
        left:${Math.random()*100}vw;top:100vh;
        animation:confetti-up 1.6s ease-out ${i*45}ms forwards;
      `;
      document.body.appendChild(dot);
      setTimeout(() => dot.remove(), 2200);
    }
  }

  return { init };
})();

/* ══════════ 13. BACK TO TOP ══════════ */
document.getElementById('backToTop')?.addEventListener('click', () =>
  window.scrollTo({ top: 0, behavior: 'smooth' })
);

/* ══════════ 14. DYNAMIC CSS ══════════ */
(function injectCSS() {
  const s = document.createElement('style');
  s.textContent = `
    @keyframes confetti-up {
      0%   { transform:translateY(0) rotate(0deg); opacity:1; }
      100% { transform:translateY(-100vh) rotate(720deg); opacity:0; }
    }
  `;
  document.head.appendChild(s);
})();

/* ══════════ INIT ALL ══════════ */
document.addEventListener('DOMContentLoaded', () => {
  Theme.init();
  Nav.init();
  AOS.init();
  Counter.init();
  Tabs.init();
  Tilt.init();
  Typewriter.init();
  Form.init();
  setTimeout(() => Particles.init(), 120);
  window.dispatchEvent(new Event('scroll'));
});