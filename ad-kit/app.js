(function () {
  const root = document.getElementById("slideshow");
  const headlineEl = document.getElementById("sceneHeadline");
  const sublineEl = document.getElementById("sceneSubline");
  const tagEl = document.getElementById("sceneTag");
  const benefitEls = [
    document.getElementById("benefit1"),
    document.getElementById("benefit2"),
    document.getElementById("benefit3"),
  ];
  const progressBar = document.getElementById("progressBar");
  const narrationEl = document.getElementById("narration");

  if (!root) return;

  const slidePaths = [
    "assets/shot1.png",
    "assets/shot2.png",
    "assets/shot3.png",
    "assets/shot4.png",
    "assets/shot5.png",
  ];

  const scenes = [
    {
      tag: "Problem",
      headline: "Canvas is wasting your study time.",
      subline: "Deadlines are buried and tabs get out of control.",
      narration:
        "If Canvas feels cluttered, QuickCanvas gives you one clean place to see what matters first.",
      benefits: [
        "Stop hunting for what is due next",
        "Cut clutter with cleaner views",
        "Keep focus on actual homework",
      ],
    },
    {
      tag: "Fix",
      headline: "QuickCanvas reorganizes your dashboard.",
      subline: "Your most urgent tasks move to the front instantly.",
      narration:
        "It adds due-date filters and smart dashboard sections so assignments stop getting lost.",
      benefits: [
        "Next due date filtering in one tap",
        "Personal to-do right inside Canvas",
        "Smart panels that reduce scroll fatigue",
      ],
    },
    {
      tag: "Style",
      headline: "Make Canvas feel premium, not outdated.",
      subline: "Dark themes and cleaner contrast that actually look good.",
      narration:
        "You can customize colors, fonts, and layout so Canvas is easier to read and easier to use.",
      benefits: [
        "Custom accents and fonts",
        "Readable cards and cleaner spacing",
        "Setup takes under a minute",
      ],
    },
    {
      tag: "Proof",
      headline: "Built for real student workflow.",
      subline: "Less chaos, fewer missed assignments, faster study starts.",
      narration:
        "This is made for busy students who want less admin stress and faster homework starts.",
      benefits: [
        "Open extension, pick your layout",
        "See due dates before they surprise you",
        "Stay organized through heavy weeks",
      ],
    },
    {
      tag: "Try it",
      headline: "Ready to make Canvas way easier?",
      subline: "Install QuickCanvas and upgrade your dashboard today.",
      narration:
        "Install QuickCanvas from the Chrome Web Store and upgrade your Canvas setup in minutes.",
      benefits: ["Fast setup", "Clean design", "Student-first workflow"],
    },
    {
      tag: "Result",
      headline: "Spend less time organizing, more time studying.",
      subline: "This is what Canvas should have felt like from day one.",
      narration:
        "If you use Canvas daily, this extension gives you a cleaner system to stay on track.",
      benefits: [
        "Clear due-date visibility",
        "Cleaner dashboard flow",
        "More focused work sessions",
      ],
    },
  ];

  const SLIDE_MS = 3400;
  const TRANSITION_MS = 500;

  function loadImage(path) {
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => resolve({ ok: true, path });
      img.onerror = () => resolve({ ok: false, path });
      img.src = path;
    });
  }

  function createSlide(path) {
    const slide = document.createElement("div");
    slide.className = "slide";
    const img = document.createElement("img");
    img.src = path;
    img.alt = "QuickCanvas screenshot";
    slide.appendChild(img);
    return slide;
  }

  function createFallbackSlide() {
    const slide = document.createElement("div");
    slide.className = "slide fallback active";
    const text = document.createElement("p");
    text.textContent =
      "Add screenshots to assets/shot1.png ... assets/shot5.png to preview this ad.";
    slide.appendChild(text);
    return slide;
  }

  function applyScene(index, total) {
    const scene = scenes[index % scenes.length];
    if (tagEl) tagEl.textContent = scene.tag;
    if (headlineEl) headlineEl.textContent = scene.headline;
    if (sublineEl) sublineEl.textContent = scene.subline;
    if (narrationEl) narrationEl.textContent = scene.narration;
    scene.benefits.forEach((text, idx) => {
      if (!benefitEls[idx]) return;
      benefitEls[idx].textContent = text;
      benefitEls[idx].style.opacity = idx === 0 ? "1" : "0.95";
    });
    if (progressBar) {
      const percent = ((index + 1) / Math.max(total, 1)) * 100;
      progressBar.style.width = percent + "%";
    }
  }

  async function init() {
    const checks = await Promise.all(slidePaths.map((path) => loadImage(path)));
    const valid = checks.filter((item) => item.ok).map((item) => item.path);

    if (valid.length === 0) {
      root.appendChild(createFallbackSlide());
      applyScene(0, 1);
      return;
    }

    const slides = valid.map(createSlide);
    slides.forEach((slide) => root.appendChild(slide));
    slides[0].classList.add("active");
    applyScene(0, slides.length);

    let current = 0;
    setInterval(() => {
      const next = (current + 1) % slides.length;
      const currentSlide = slides[current];
      const nextSlide = slides[next];

      currentSlide.classList.remove("active");
      currentSlide.classList.add("exiting");
      nextSlide.classList.add("active");

      setTimeout(() => {
        currentSlide.classList.remove("exiting");
      }, TRANSITION_MS);

      current = next;
      applyScene(current, slides.length);
    }, SLIDE_MS);
  }

  init();
})();
