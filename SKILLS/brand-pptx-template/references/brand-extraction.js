// Browser console snippet to extract colors/fonts/logo from a website.
(() => {
  const colors = new Map();
  const fonts = new Set();

  const selectors = [
    'body', 'header', 'nav', 'footer',
    'h1', 'h2', 'h3', 'p', 'a',
    'button', '[class*="btn"]', '[class*="card"]',
    '[class*="hero"]', '[class*="cta"]'
  ];

  selectors.forEach((sel) => {
    document.querySelectorAll(sel).forEach((el) => {
      const style = getComputedStyle(el);
      ['color', 'backgroundColor', 'borderColor'].forEach((prop) => {
        const val = style[prop];
        if (val && val !== 'rgba(0, 0, 0, 0)' && val !== 'transparent') {
          colors.set(val, (colors.get(val) || 0) + 1);
        }
      });

      const font = style.fontFamily
        .split(',')[0]
        .trim()
        .replace(/["']/g, '');
      if (font) fonts.add(font);
    });
  });

  const logo = document.querySelector(
    'img[src*="logo"], header img, .logo img, [class*="brand"] img'
  );

  return {
    topColors: [...colors.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, 15),
    fonts: [...fonts],
    logoSrc: logo?.src || null,
  };
})();
