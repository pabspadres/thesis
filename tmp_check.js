
    const controlFile = "./Data/Control_Database.xlsx";
    const colorFile = "./Data/Color_Database.xlsx";
    const scatterTicks = [0, 25, 50, 75, 100];
    const barsTicks = [0, 50, 100, 150, 200];
    const introTokenDefs = [
      { id: "lululemon", file: "./media/lululemon.png", brands: ["lululemon"] },
      { id: "levi", file: "./media/levi.png", brands: ["levi strauss & co.", "levi strauss and co", "levi"] },
      { id: "rl", file: "./media/RL.png", brands: ["ralph lauren"] },
      { id: "gap", file: "./media/gap.png", brands: ["gap"] },
      { id: "nike", file: "./media/nike.png", brands: ["nike"] },
      { id: "patagonia", file: "./media/patagonia.png", brands: ["patagonia"] }
    ];
    const brandLogoMap = new Map([
      ["gap", "./media/gap.png"],
      ["abercrombie & fitch", "./media/A&F.png"],
      ["american eagle", "./media/AE.png"],
      ["american eagle outfitters", "./media/AE.png"],
      ["anthropologie", "./media/anthropologie.png"],
      ["banana republic", "./media/BR.png"],
      ["champion", "./media/champion.png"],
      ["calvin klein", "./media/CK.png"],
      ["clavin klein", "./media/CK.png"],
      ["calvin klein inc.", "./media/CK.png"],
      ["calvin klein, inc.", "./media/CK.png"],
      ["coach", "./media/coach.png"],
      ["coach new york", "./media/coach.png"],
      ["converse", "./media/converse.png"],
      ["fruit of the loom", "./media/FL.png"],
      ["forever 21", "./media/forever21.png"],
      ["hanes", "./media/hanes.png"],
      ["kate spade & co", "./media/KS.png"],
      ["kate spade & co.", "./media/KS.png"],
      ["kate spade & company", "./media/KS.png"],
      ["kate spade", "./media/KS.png"],
      ["levi strauss & co", "./media/levi.png"],
      ["levi strauss & co.", "./media/levi.png"],
      ["levi strauss and co", "./media/levi.png"],
      ["ll. bean", "./media/llbean.png"],
      ["ll bean", "./media/llbean.png"],
      ["lululemon", "./media/lululemon.png"],
      ["lululemon athletica", "./media/lululemon.png"],
      ["marc jacobs", "./media/MJ.png"],
      ["michael kors", "./media/MK.png"],
      ["new balance", "./media/NB.png"],
      ["the north face", "./media/NF.png"],
      ["the north face, inc.", "./media/NF.png"],
      ["north face", "./media/NF.png"],
      ["nike", "./media/nike.png"],
      ["old navy", "./media/ON.png"],
      ["patagonia", "./media/patagonia.png"],
      ["ralph lauren corporation", "./media/RL.png"],
      ["ralph lauren", "./media/RL.png"],
      ["steve madden", "./media/SM.png"],
      ["tory burch", "./media/TB.png"],
      ["tom ford", "./media/TF.png"],
      ["tommy hilfiger", "./media/TH.png"],
      ["under armour", "./media/UA.png"],
      ["urban outfitters", "./media/UB.png"],
      ["victoria’s secret & co", "./media/VS.png"],
      ["victoria's secret & co", "./media/VS.png"],
      ["victoria's secret", "./media/VS.png"],
      ["victoria’s secret", "./media/VS.png"],
      ["wrangler", "./media/wrangler.png"]
    ]);

    function normalizeBrand(value) {
      return String(value || "").trim().toLowerCase();
    }

    function readWorkbookFromBuffer(buffer) {
      return XLSX.read(buffer, { type: "array" });
    }

    function getFirstSheetRows(workbook) {
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      return XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: true });
    }

    function findHeaderRow(rows, requiredHeaders) {
      const required = requiredHeaders.map((h) => h.toLowerCase());
      return rows.findIndex((row) => {
        const rowHeaders = row
          .map((cell) => (typeof cell === "string" ? cell.trim().toLowerCase() : ""))
          .filter(Boolean);
        return required.every((requiredHeader) => rowHeaders.includes(requiredHeader));
      });
    }

    function parseControlRows(rows) {
      const headerRowIndex = findHeaderRow(rows, ["Brand", "Ethical Compliance Score", "Scale Score (Revenue)", "Final Score"]);
      if (headerRowIndex === -1) {
        throw new Error("Control_Database.xlsx header row not found.");
      }

      const headers = rows[headerRowIndex].map((v) => (typeof v === "string" ? v.trim() : v));
      const brandIdx = headers.findIndex((h) => String(h).toLowerCase() === "brand");
      const ethicalIdx = headers.findIndex((h) => String(h).toLowerCase() === "ethical compliance score");
      const scaleIdx = headers.findIndex((h) => String(h).toLowerCase() === "scale score (revenue)");
      const finalIdx = headers.findIndex((h) => String(h).toLowerCase() === "final score");

      const output = [];
      for (let i = headerRowIndex + 1; i < rows.length; i += 1) {
        const row = rows[i];
        if (!row || row.length === 0) {
          continue;
        }

        const brand = row[brandIdx];
        const ethicalScore = Number(row[ethicalIdx]);
        const scaleScore = Number(row[scaleIdx]);
        const finalScore = Number(row[finalIdx]);

        if (!brand || Number.isNaN(ethicalScore) || Number.isNaN(scaleScore) || Number.isNaN(finalScore)) {
          continue;
        }

        output.push({
          brand: String(brand).trim(),
          ethical: Math.max(0, Math.min(100, ethicalScore)),
          scale: Math.max(0, Math.min(100, scaleScore)),
          final: Math.max(0, Math.min(200, finalScore))
        });
      }
      return output;
    }

    function parseColorRows(rows) {
      const headerRowIndex = findHeaderRow(rows, ["Brand", "Color"]);
      if (headerRowIndex === -1) {
        throw new Error("Color_Database.xlsx header row not found.");
      }

      const headers = rows[headerRowIndex].map((v) => (typeof v === "string" ? v.trim() : v));
      const brandIdx = headers.findIndex((h) => String(h).toLowerCase() === "brand");
      const colorIdx = headers.findIndex((h) => String(h).toLowerCase() === "color");

      const colorMap = new Map();
      for (let i = headerRowIndex + 1; i < rows.length; i += 1) {
        const row = rows[i];
        if (!row || row.length === 0) {
          continue;
        }

        const brand = row[brandIdx];
        const color = row[colorIdx];

        if (!brand || !color) {
          continue;
        }

        colorMap.set(normalizeBrand(brand), String(color).trim());
      }
      return colorMap;
    }

    function svgText(content, x, y, extraAttrs = "") {
      return `<text x="${x}" y="${y}" ${extraAttrs}>${content}</text>`;
    }

    function escapeXml(value) {
      return String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\"/g, "&quot;")
        .replace(/'/g, "&apos;");
    }

    function getScatterLayout(svgElement) {
      const width = Math.max(640, svgElement.clientWidth);
      const height = Math.max(420, svgElement.clientHeight);
      const margin = { top: 24, right: 30, bottom: 72, left: 86 };
      const chartW = width - margin.left - margin.right;
      const chartH = height - margin.top - margin.bottom;
      const xToPx = (value) => margin.left + (value / 100) * chartW;
      const yToPx = (value) => margin.top + chartH - (value / 100) * chartH;

      return {
        width,
        height,
        margin,
        chartW,
        chartH,
        xToPx,
        yToPx
      };
    }

    function buildScatterAxisAndTicks(layout) {
      const { margin, chartW, chartH, height, xToPx, yToPx } = layout;
      const axisAndTicks = [];

      axisAndTicks.push(`<line x1="${margin.left}" y1="${margin.top + chartH}" x2="${margin.left + chartW}" y2="${margin.top + chartH}" stroke="#000" stroke-width="1.5" />`);
      axisAndTicks.push(`<line x1="${margin.left}" y1="${margin.top + chartH}" x2="${margin.left}" y2="${margin.top}" stroke="#000" stroke-width="1.5" />`);

      scatterTicks.forEach((tick) => {
        const x = xToPx(tick);
        const y = yToPx(tick);

        axisAndTicks.push(`<line x1="${x}" y1="${margin.top + chartH}" x2="${x}" y2="${margin.top + chartH + 6}" stroke="#000" stroke-width="1" />`);
        axisAndTicks.push(svgText(tick, x, margin.top + chartH + 24, 'text-anchor="middle" font-family="MontserratCustom" font-size="13" fill="#000" font-weight="400"'));

        axisAndTicks.push(`<line x1="${margin.left - 6}" y1="${y}" x2="${margin.left}" y2="${y}" stroke="#000" stroke-width="1" />`);
        axisAndTicks.push(svgText(tick, margin.left - 12, y + 4, 'text-anchor="end" font-family="MontserratCustom" font-size="13" fill="#000" font-weight="400"'));
      });

      axisAndTicks.push(svgText("Ethical Compliance Score", margin.left + chartW / 2, height - 22, 'text-anchor="middle" font-family="MontserratCustom" font-size="15" fill="#000" font-weight="400"'));
      axisAndTicks.push(`<text x="24" y="${margin.top + chartH / 2}" transform="rotate(-90 24 ${margin.top + chartH / 2})" text-anchor="middle" font-family="MontserratCustom" font-size="15" fill="#000" font-weight="400">Revenue Scale Score</text>`);

      return axisAndTicks;
    }


    function getBrandLogoForScatter(brandName) {
      const normalized = normalizeBrand(brandName);
      if (brandLogoMap.has(normalized)) {
        return brandLogoMap.get(normalized);
      }
      return null;
    }

    function renderLogoScatter(points) {
      const svg = document.getElementById("insights-chart");
      if (!svg) {
        return;
      }
      const layout = getScatterLayout(svg);
      const { width, height, xToPx, yToPx } = layout;
      const axisAndTicks = buildScatterAxisAndTicks(layout);
      const logoSize = 68;
      const mappedPoints = points.filter((point) => getBrandLogoForScatter(point.brand));
      const hiddenAll = insightsState.mode === "hidden-all";
      const visibleBrandKeys = insightsState.mode === "subset-only"
        ? new Set(insightsState.visibleBrandKeys || [])
        : null;

      const markers = mappedPoints
        .map((point) => {
          const cx = xToPx(point.ethical);
          const cy = yToPx(point.scale);
          const logoPath = getBrandLogoForScatter(point.brand);
          const brandKey = normalizeBrand(point.brand);
          const isVisible = !hiddenAll && (!visibleBrandKeys || visibleBrandKeys.has(brandKey));
          const visibilityClass = isVisible ? "is-visible" : "is-hidden";
          return `<image class="scatter-logo ${visibilityClass}" data-brand="${escapeXml(point.brand)}" data-brand-key="${escapeXml(brandKey)}" href="${logoPath}" x="${cx - logoSize / 2}" y="${cy - logoSize / 2}" width="${logoSize}" height="${logoSize}" preserveAspectRatio="xMidYMid meet" />`;
        })
        .join("");

      svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
      svg.innerHTML = `<g>${axisAndTicks.join("")}${markers}</g>
      <g id="insights-tooltip" visibility="hidden" pointer-events="none">
        <rect id="insights-tooltip-bg" x="0" y="0" width="0" height="0" rx="4" ry="4" fill="rgba(0,0,0,0.6)"></rect>
        <text id="insights-tooltip-text" x="0" y="0" fill="#fff" font-family="MontserratCustom" font-size="13" font-weight="400"></text>
      </g>`;

      const tooltipGroup = svg.querySelector("#insights-tooltip");
      const tooltipBg = svg.querySelector("#insights-tooltip-bg");
      const tooltipText = svg.querySelector("#insights-tooltip-text");
      const markerElements = svg.querySelectorAll(".scatter-logo");

      markerElements.forEach((marker) => {
        marker.addEventListener("mouseenter", () => {
          if (marker.classList.contains("is-hidden")) {
            return;
          }
          const brand = marker.getAttribute("data-brand") || "";
          tooltipText.textContent = brand;
          const textBox = tooltipText.getBBox();
          const paddingX = 8;
          const paddingY = 6;
          const boxWidth = textBox.width + paddingX * 2;
          const boxHeight = textBox.height + paddingY * 2;

          const x = Number(marker.getAttribute("x"));
          const y = Number(marker.getAttribute("y"));
          const w = Number(marker.getAttribute("width"));
          const h = Number(marker.getAttribute("height"));
          const cx = x + w / 2;
          const cy = y + h / 2;

          let boxX = cx + 12;
          let boxY = cy - boxHeight - 12;
          if (boxX + boxWidth > width - 4) {
            boxX = cx - boxWidth - 12;
          }
          if (boxY < 4) {
            boxY = cy + 12;
          }

          tooltipBg.setAttribute("x", String(boxX));
          tooltipBg.setAttribute("y", String(boxY));
          tooltipBg.setAttribute("width", String(boxWidth));
          tooltipBg.setAttribute("height", String(boxHeight));
          tooltipText.setAttribute("x", String(boxX + paddingX));
          tooltipText.setAttribute("y", String(boxY + boxHeight - paddingY - 1));
          tooltipGroup.setAttribute("visibility", "visible");
        });

        marker.addEventListener("mouseleave", () => {
          tooltipGroup.setAttribute("visibility", "hidden");
        });
      });
    }

    function getIntroTransitionBrandKeys() {
      const keys = [];
      introTokenDefs.forEach((tokenDef) => {
        const point = getIntroBrandPoint(tokenDef);
        if (point) {
          keys.push(normalizeBrand(point.brand));
        }
      });
      return keys;
    }

    function setInsightsMode(mode, visibleBrandKeys = []) {
      insightsState.mode = mode;
      insightsState.visibleBrandKeys = visibleBrandKeys;
      if (currentPoints.length > 0) {
        renderLogoScatter(currentPoints);
      }
    }

    function revealRemainingInsightsLogos() {
      const brandKeysToKeep = new Set(getIntroTransitionBrandKeys());
      const logoElements = Array.from(document.querySelectorAll("#insights-chart .scatter-logo"));
      const extraLogos = logoElements.filter((logo) => !brandKeysToKeep.has(logo.getAttribute("data-brand-key") || ""));

      extraLogos.forEach((logo, index) => {
        window.setTimeout(() => {
          logo.classList.remove("is-hidden");
          logo.classList.add("is-visible");
        }, index * 45);
      });
    }

    function renderIntroChart() {
      const svg = document.getElementById("intro-chart");
      const layout = getScatterLayout(svg);
      const axisAndTicks = buildScatterAxisAndTicks(layout);
      svg.setAttribute("viewBox", `0 0 ${layout.width} ${layout.height}`);
      svg.innerHTML = `<g>${axisAndTicks.join("")}</g>`;
    }

    const introState = {
      tokenMap: new Map(),
      dragId: null,
      dragOffsetX: 0,
      dragOffsetY: 0,
      revealTimer: null,
      transitionTimer: null,
      revealed: false,
      transitioningToInsights: false
    };
    const insightsState = {
      mode: "all-visible"
    };
    const methodIntroState = {
      revealed: false,
      transitioningToBag: false,
      waitingForSecondScrollToBag: false,
      entryLockUntilMs: 0,
      suppressMethodDownUntilIdle: false,
      methodDownIdleTimer: null,
      lastWheelEventAtMs: 0,
      wheelGestureId: 0,
      revealGestureId: -1
    };
    const metricsBagState = {
      canScrollToGameIntro: false
    };
    const metricDetails = {
      "metric-vde": {
        title: "Value Distribution Equity",
        text: "Fill with sample text",
        placement: "right"
      },
      "metric-lc": {
        title: "Labor Compensation",
        text: "Fill with sample text",
        placement: "right"
      },
      "metric-st": {
        title: "Social Transparency",
        text: "This metric extracts scores from the Fashion Transparency Index by Fashion Revolution, which provides scores from 2023, grading the brand on their public disclosure on social and environmental impact.",
        placement: "bottom"
      },
      "metric-ei": {
        title: "Environmental Impact",
        text: "This metric extracts scores from the What Fuels Fashion analysis by Fashion Revolution, which provides scores from 2025 grading the brand on their disclosure of their climate and energy-related policies, practices and impacts in their operations.",
        placement: "left"
      },
      "metric-sct": {
        title: "Supply Chain Transparency",
        text: "This metric extracts  scores from a dataset from the Supply Chain Transparency Score from the The Clean Clothes Campaign through Wikirate. These scores rate brands based on how transparent they are about their suppliers.",
        placement: "left"
      },
      "metric-hrrr": {
        title: "Human Rights Risk & Responsiveness",
        text: "This scores for this metric rate a brand based on the data from the Business & Human Rights Centre, which reports on news and allegations relating to the human rights impact of brands and their response if any. The scores were weighted and calculated using machine learning.",
        placement: "top"
      }
    };

    function getIntroTokenById(tokenId) {
      return introState.tokenMap.get(tokenId);
    }

    function setTokenPosition(tokenState, x, y) {
      tokenState.x = x;
      tokenState.y = y;
      tokenState.el.style.left = `${x}px`;
      tokenState.el.style.top = `${y}px`;
    }

    function getIntroBrandPoint(tokenDef) {
      for (let i = 0; i < tokenDef.brands.length; i += 1) {
        const targetBrand = normalizeBrand(tokenDef.brands[i]);
        const exact = currentPoints.find((point) => normalizeBrand(point.brand) === targetBrand);
        if (exact) {
          return exact;
        }
      }

      for (let i = 0; i < tokenDef.brands.length; i += 1) {
        const targetBrand = normalizeBrand(tokenDef.brands[i]);
        const partial = currentPoints.find((point) => normalizeBrand(point.brand).includes(targetBrand));
        if (partial) {
          return partial;
        }
      }

      return null;
    }

    function placeIntroTokensInTray() {
      const content = document.getElementById("intro-content");
      const logoRow = document.getElementById("intro-logo-row");
      const contentRect = content.getBoundingClientRect();
      const logoRect = logoRow.getBoundingClientRect();

      const tokenHeight = Math.max(74, logoRect.height - 18);
      const tokenWidth = tokenHeight * (124 / 74);
      const usableWidth = logoRect.width - tokenWidth;
      const startX = logoRect.left - contentRect.left;
      const y = logoRect.top - contentRect.top + (logoRect.height - tokenHeight) / 2;

      introTokenDefs.forEach((tokenDef, index) => {
        const tokenState = getIntroTokenById(tokenDef.id);
        if (!tokenState || tokenState.placed) {
          return;
        }

        const sizeBoost = (tokenDef.id === "patagonia" || tokenDef.id === "lululemon") ? 1.2 : 1;
        const currentTokenWidth = tokenWidth * sizeBoost;
        const currentTokenHeight = tokenHeight * sizeBoost;
        tokenState.el.style.width = `${tokenWidth}px`;
        tokenState.el.style.height = `${tokenHeight}px`;
        tokenState.el.style.width = `${currentTokenWidth}px`;
        tokenState.el.style.height = `${currentTokenHeight}px`;

        const progress = introTokenDefs.length > 1 ? index / (introTokenDefs.length - 1) : 0;
        const x = startX + usableWidth * progress - (currentTokenWidth - tokenWidth) / 2;
        const currentY = y - (currentTokenHeight - tokenHeight) / 2;
        setTokenPosition(tokenState, x, currentY);
      });
    }

    function maybeStartRevealTimer() {
      if (introState.revealed || introState.revealTimer) {
        return;
      }

      const allPlaced = introTokenDefs.every((tokenDef) => {
        const tokenState = getIntroTokenById(tokenDef.id);
        return tokenState && tokenState.placed;
      });

      if (!allPlaced) {
        return;
      }

      introState.revealTimer = window.setTimeout(() => {
        introState.revealTimer = null;
        revealIntroCorrectPlacements();
      }, 1000);
    }

    function revealIntroCorrectPlacements(force = false) {
      if (introState.revealed && !force) {
        return;
      }

      const content = document.getElementById("intro-content");
      const chart = document.getElementById("intro-chart");
      const contentRect = content.getBoundingClientRect();
      const chartRect = chart.getBoundingClientRect();
      const layout = getScatterLayout(chart);

      introTokenDefs.forEach((tokenDef) => {
        const tokenState = getIntroTokenById(tokenDef.id);
        const point = getIntroBrandPoint(tokenDef);
        if (!tokenState || !point) {
          return;
        }

        const targetX = chartRect.left - contentRect.left + layout.xToPx(point.ethical) - tokenState.el.offsetWidth / 2;
        const targetY = chartRect.top - contentRect.top + layout.yToPx(point.scale) - tokenState.el.offsetHeight / 2;
        tokenState.placed = true;
        setTokenPosition(tokenState, targetX, targetY);
      });

      if (!introState.revealed) {
        introState.revealed = true;
        if (introState.transitionTimer) {
          window.clearTimeout(introState.transitionTimer);
        }
        introState.transitionTimer = window.setTimeout(() => {
          startGameToInsightsTransition();
        }, 2000);
      }
    }

    function introOnPointerMove(event) {
      if (!introState.dragId) {
        return;
      }

      const content = document.getElementById("intro-content");
      const contentRect = content.getBoundingClientRect();
      const tokenState = getIntroTokenById(introState.dragId);
      if (!tokenState) {
        return;
      }

      const x = event.clientX - contentRect.left - introState.dragOffsetX;
      const y = event.clientY - contentRect.top - introState.dragOffsetY;
      setTokenPosition(tokenState, x, y);
    }

    function introOnPointerUp() {
      if (!introState.dragId) {
        return;
      }

      const tokenState = getIntroTokenById(introState.dragId);
      const content = document.getElementById("intro-content");
      const chart = document.getElementById("intro-chart");
      const contentRect = content.getBoundingClientRect();
      const chartRect = chart.getBoundingClientRect();

      if (tokenState) {
        tokenState.el.classList.remove("is-dragging");
        tokenState.el.style.zIndex = "20";
        const centerX = contentRect.left + tokenState.x + tokenState.el.offsetWidth / 2;
        const centerY = contentRect.top + tokenState.y + tokenState.el.offsetHeight / 2;
        tokenState.placed = (
          centerX >= chartRect.left
          && centerX <= chartRect.right
          && centerY >= chartRect.top
          && centerY <= chartRect.bottom
        );
      }

      introState.dragId = null;

      if (introState.revealTimer) {
        const allPlaced = introTokenDefs.every((tokenDef) => {
          const state = getIntroTokenById(tokenDef.id);
          return state && state.placed;
        });
        if (!allPlaced) {
          window.clearTimeout(introState.revealTimer);
          introState.revealTimer = null;
        }
      }

      maybeStartRevealTimer();
    }

    function startIntroDrag(event, tokenDef) {
      const tokenState = getIntroTokenById(tokenDef.id);
      if (!tokenState || introState.revealed) {
        return;
      }

      const tokenRect = tokenState.el.getBoundingClientRect();
      introState.dragId = tokenDef.id;
      introState.dragOffsetX = event.clientX - tokenRect.left;
      introState.dragOffsetY = event.clientY - tokenRect.top;
      tokenState.el.classList.add("is-dragging");
      tokenState.el.style.zIndex = "50";
    }

    function createIntroTokens() {
      const tokenLayer = document.getElementById("intro-token-layer");
      tokenLayer.innerHTML = "";
      introState.tokenMap.clear();

      introTokenDefs.forEach((tokenDef) => {
        const token = document.createElement("div");
        token.className = "intro-token";
        token.dataset.tokenId = tokenDef.id;

        const image = document.createElement("img");
        image.src = tokenDef.file;
        image.alt = tokenDef.id;
        token.appendChild(image);

        token.addEventListener("pointerdown", (event) => {
          startIntroDrag(event, tokenDef);
        });

        tokenLayer.appendChild(token);
        introState.tokenMap.set(tokenDef.id, {
          id: tokenDef.id,
          el: token,
          x: 0,
          y: 0,
          placed: false
        });
      });

      document.addEventListener("pointermove", introOnPointerMove);
      document.addEventListener("pointerup", introOnPointerUp);
      placeIntroTokensInTray();
    }

    function showMainInterface() {
      const introView = document.getElementById("intro-view");
      const insightsPage = document.getElementById("insights-page-1");

      if (introState.revealTimer) {
        window.clearTimeout(introState.revealTimer);
        introState.revealTimer = null;
      }
      if (introState.transitionTimer) {
        window.clearTimeout(introState.transitionTimer);
        introState.transitionTimer = null;
      }

      if (introView) {
        introView.style.display = "none";
      }

      if (insightsPage) {
        setInsightsMode("all-visible");
        insightsPage.classList.add("is-visible");
        document.documentElement.classList.add("no-snap");
        document.body.classList.add("no-snap");
        insightsPage.scrollIntoView({ behavior: "auto", block: "start" });
        window.requestAnimationFrame(() => {
          insightsPage.scrollIntoView({ behavior: "auto", block: "start" });
        });
      }
    }

    function startGameToInsightsTransition() {
      if (introState.transitioningToInsights) {
        return;
      }

      const introView = document.getElementById("intro-view");
      const insightsPage = document.getElementById("insights-page-1");
      const insightsChart = document.getElementById("insights-chart");
      const overlay = document.getElementById("transition-overlay");
      if (!introView || !insightsPage || !insightsChart || !overlay) {
        showMainInterface();
        return;
      }

      introState.transitioningToInsights = true;
      const transitionBrandKeys = new Set(getIntroTransitionBrandKeys());
      setInsightsMode("hidden-all");
      insightsPage.classList.add("is-visible");
      document.documentElement.classList.add("no-snap");
      document.body.classList.add("no-snap");
      // Keep current frame visible while we re-parent the existing tokens.

      window.requestAnimationFrame(() => {
        const tokensToAnimate = [];
        overlay.innerHTML = "";

        // Move the EXISTING 6 token elements into the overlay as fixed-position,
        // so they can shrink smoothly without ever disappearing.
        introTokenDefs.forEach((tokenDef) => {
          const tokenState = getIntroTokenById(tokenDef.id);
          const point = getIntroBrandPoint(tokenDef);
          if (!tokenState || !point) {
            return;
          }

          const rect = tokenState.el.getBoundingClientRect();
          tokenState.el.classList.remove("is-dragging");
          tokenState.el.style.transition = "none";
          tokenState.el.style.position = "fixed";
          tokenState.el.style.left = `${rect.left}px`;
          tokenState.el.style.top = `${rect.top}px`;
          tokenState.el.style.width = `${rect.width}px`;
          tokenState.el.style.height = `${rect.height}px`;
          tokenState.el.style.margin = "0";
          tokenState.el.style.zIndex = "202";
          tokenState.el.style.cursor = "default";
          tokenState.el.style.pointerEvents = "none";
          tokenState.el.style.opacity = "1";
          overlay.appendChild(tokenState.el);
          tokensToAnimate.push({ el: tokenState.el, point });
        });

        // Now jump to the ECI page while the tokens stay fixed on screen.
        insightsPage.scrollIntoView({ behavior: "auto", block: "start" });

        const layout = getScatterLayout(insightsChart);
        const chartRect = insightsChart.getBoundingClientRect();
        // Animate the same token elements into the ECI positions + ECI size.
        window.requestAnimationFrame(() => {
          tokensToAnimate.forEach(({ el, point }) => {
            el.style.transition = "left 650ms ease, top 650ms ease, width 650ms ease, height 650ms ease";
            const targetSize = 68;
            const targetLeft = chartRect.left + layout.xToPx(point.ethical) - targetSize / 2;
            const targetTop = chartRect.top + layout.yToPx(point.scale) - targetSize / 2;
            el.style.left = `${targetLeft}px`;
            el.style.top = `${targetTop}px`;
            el.style.width = `${targetSize}px`;
            el.style.height = `${targetSize}px`;
          });
        });

        window.setTimeout(() => {
          introView.style.display = "none";
          setInsightsMode("subset-only", Array.from(transitionBrandKeys));
          overlay.innerHTML = "";
          revealRemainingInsightsLogos();
          introState.transitioningToInsights = false;
        }, 700);
      });
    }

    function initScrollTransitions() {
      const gamePage = document.getElementById("game-page");
      const infoPage = document.getElementById("info-page");
      const playIntroPage = document.getElementById("play-intro-page");
      const insightsPage = document.getElementById("insights-page-1");

      const observer = new IntersectionObserver(
        (entries) => {
          entries.forEach((entry) => {
            if (entry.target.id !== "game-page") {
              if (entry.target.id === "insights-page-1") {
                insightsPage.classList.toggle("is-visible", entry.isIntersecting && entry.intersectionRatio > 0.2);
              }
              return;
            }
            if (entry.isIntersecting && entry.intersectionRatio > 0.2) {
              gamePage.classList.add("is-visible");
              infoPage.classList.add("is-fading");
              playIntroPage.classList.add("is-fading");
            } else {
              gamePage.classList.remove("is-visible");
              infoPage.classList.remove("is-fading");
              playIntroPage.classList.remove("is-fading");
            }
          });
        },
        {
          threshold: [0, 0.2, 0.5, 0.8, 1]
        }
      );

      observer.observe(gamePage);
      observer.observe(insightsPage);
    }

    function initPlayGameNavigation() {
      const methodIntroPage = document.getElementById("method-intro-page");
      const metricsBagPage = document.getElementById("metrics-bag-page");
      const metricsBagImage = document.getElementById("metrics-bag-image");
      const metricDetailBox = document.getElementById("metric-detail-box");
      const metricDetailTitle = document.getElementById("metric-detail-title");
      const metricDetailText = document.getElementById("metric-detail-text");
      const metricItems = Array.from(document.querySelectorAll(".metrics-burst-item"));
      const playIntroPage = document.getElementById("play-intro-page");
      const gamePage = document.getElementById("game-page");
      const insightsPage = document.getElementById("insights-page-1");
      const playButton = document.getElementById("play-button");
      const playSkipButton = document.getElementById("play-skip-button");
      if (!methodIntroPage || !metricsBagPage || !metricsBagImage || !metricDetailBox || !metricDetailTitle || !metricDetailText || !playIntroPage || !gamePage || !insightsPage || !playButton || !playSkipButton) {
        return;
      }

      methodIntroPage.classList.add("is-pre-reveal");

      const resetMethodIntroReveal = () => {
        methodIntroState.revealed = false;
        methodIntroState.transitioningToBag = false;
        methodIntroState.waitingForSecondScrollToBag = false;
        methodIntroState.entryLockUntilMs = 0;
        methodIntroState.suppressMethodDownUntilIdle = false;
        methodIntroState.lastWheelEventAtMs = 0;
        methodIntroState.wheelGestureId = 0;
        methodIntroState.revealGestureId = -1;
        if (methodIntroState.methodDownIdleTimer) {
          window.clearTimeout(methodIntroState.methodDownIdleTimer);
          methodIntroState.methodDownIdleTimer = null;
        }
        methodIntroPage.classList.remove("is-transitioning-forward");
        methodIntroPage.classList.add("is-pre-reveal");
        const overlay = document.getElementById("transition-overlay");
        if (overlay) {
          overlay.innerHTML = "";
        }
      };

      const methodIntroObserver = new IntersectionObserver(
        (entries) => {
          entries.forEach((entry) => {
            if (entry.target !== methodIntroPage) {
              return;
            }
            if (entry.intersectionRatio > 0.6) {
              // Ignore carry-over wheel momentum right after snapping into this section.
              methodIntroState.entryLockUntilMs = Date.now() + 450;
              methodIntroState.suppressMethodDownUntilIdle = true;
              if (methodIntroState.methodDownIdleTimer) {
                window.clearTimeout(methodIntroState.methodDownIdleTimer);
                methodIntroState.methodDownIdleTimer = null;
              }
            }
            // When this section is mostly scrolled off, re-arm slide-in for the next time we enter
            if (entry.intersectionRatio < 0.38) {
              resetMethodIntroReveal();
            }
          });
        },
        { threshold: [0, 0.15, 0.25, 0.35, 0.38, 0.5, 0.65, 0.8, 1] }
      );
      methodIntroObserver.observe(methodIntroPage);
      const metricsBagObserver = new IntersectionObserver(
        (entries) => {
          entries.forEach((entry) => {
            if (entry.target !== metricsBagPage) {
              return;
            }
            const isActive = entry.intersectionRatio > 0.6;
            metricsBagPage.classList.toggle("is-active", isActive);
            if (!isActive) {
              metricsBagPage.classList.remove("is-opened");
              metricsBagPage.classList.remove("is-detail-open");
              metricsBagPage.classList.remove("is-scroll-ready");
              metricsBagState.canScrollToGameIntro = false;
              metricDetailBox.classList.remove("is-visible", "align-left", "align-right");
            }
          });
        },
        { threshold: [0, 0.3, 0.6, 0.85, 1] }
      );
      metricsBagObserver.observe(metricsBagPage);

      const clamp = (value, min, max) => Math.max(min, Math.min(max, value));
      const hideMetricDetail = (unlockScroll = false) => {
        metricsBagPage.classList.remove("is-detail-open");
        metricDetailBox.classList.remove("is-visible", "align-left", "align-right");
        metricItems.forEach((item) => item.classList.remove("is-focused"));
        if (unlockScroll) {
          metricsBagState.canScrollToGameIntro = true;
          metricsBagPage.classList.add("is-scroll-ready");
        }
      };

      const showMetricDetailFor = (metricElement) => {
        const detail = metricDetails[metricElement.id];
        if (!detail) {
          return;
        }

        metricDetailTitle.textContent = detail.title;
        metricDetailText.textContent = detail.text;
        metricDetailBox.classList.remove("align-left", "align-right");
        metricDetailBox.classList.add(detail.placement === "right" ? "align-right" : "align-left");
        metricDetailBox.style.left = "50px";
        metricDetailBox.style.top = "50px";
        metricDetailBox.classList.add("is-visible");

        const pageRect = metricsBagPage.getBoundingClientRect();
        const metricRect = metricElement.getBoundingClientRect();
        const boxRect = metricDetailBox.getBoundingClientRect();
        const margin = 50;
        const gap = 20;
        let left = 0;
        let top = 0;

        if (detail.placement === "top") {
          left = metricRect.left - pageRect.left;
          top = metricRect.bottom - pageRect.top + gap;
        } else if (detail.placement === "bottom") {
          left = metricRect.left - pageRect.left;
          top = metricRect.top - pageRect.top - boxRect.height - gap;
        } else if (detail.placement === "left") {
          left = metricRect.right - pageRect.left + gap;
          top = metricRect.top - pageRect.top + (metricRect.height - boxRect.height) / 2;
        } else {
          left = metricRect.left - pageRect.left - boxRect.width - gap;
          top = metricRect.top - pageRect.top + (metricRect.height - boxRect.height) / 2;
        }

        left = clamp(left, margin, window.innerWidth - boxRect.width - margin);
        top = clamp(top, margin, window.innerHeight - boxRect.height - margin);
        metricDetailBox.style.left = `${left}px`;
        metricDetailBox.style.top = `${top}px`;
        metricItems.forEach((item) => item.classList.toggle("is-focused", item === metricElement));
        metricsBagPage.classList.add("is-detail-open");
      };

      metricItems.forEach((item) => {
        item.addEventListener("click", (event) => {
          if (!metricsBagPage.classList.contains("is-opened")) {
            return;
          }
          if (metricsBagPage.classList.contains("is-detail-open")) {
            event.stopPropagation();
            hideMetricDetail(true);
            return;
          }
          event.stopPropagation();
          showMetricDetailFor(item);
        });
      });

      metricsBagPage.addEventListener("click", (event) => {
        if (!metricsBagPage.classList.contains("is-detail-open")) {
          return;
        }
        const target = event.target;
        if (target instanceof Element && target.closest(".metrics-burst-item")) {
          return;
        }
        hideMetricDetail(true);
      });

      metricsBagImage.addEventListener("click", () => {
        if (!isMetricsBagActive()) {
          return;
        }
        metricsBagPage.classList.add("is-opened");
      });

      let allowGameJump = false;
      const revealMethodIntro = () => {
        if (methodIntroState.revealed) {
          return;
        }
        methodIntroState.revealed = true;
        methodIntroState.waitingForSecondScrollToBag = true;
        methodIntroState.revealGestureId = methodIntroState.wheelGestureId;
        methodIntroPage.classList.remove("is-pre-reveal");
      };

      const startMethodToBagTransition = () => {
        if (methodIntroState.transitioningToBag) {
          return;
        }
        const methodImage = document.getElementById("method-index-image");
        const overlay = document.getElementById("transition-overlay");
        if (!methodImage || !overlay) {
          metricsBagPage.scrollIntoView({ behavior: "auto", block: "start" });
          return;
        }

        methodIntroState.transitioningToBag = true;
        methodIntroPage.classList.add("is-transitioning-forward");
        metricsBagPage.classList.remove("is-active");

        const rect = methodImage.getBoundingClientRect();
        const clone = methodImage.cloneNode(true);
        clone.className = "transition-clone";
        clone.style.left = `${rect.left}px`;
        clone.style.top = `${rect.top}px`;
        clone.style.width = `${rect.width}px`;
        clone.style.height = `${rect.height}px`;
        clone.style.opacity = "1";
        overlay.innerHTML = "";
        overlay.appendChild(clone);

        const targetWidth = rect.width * 1.5;
        const targetHeight = rect.height * 1.5;
        const targetLeft = (window.innerWidth - targetWidth) / 2;
        const targetTop = (window.innerHeight - targetHeight) / 2;

        window.requestAnimationFrame(() => {
          clone.style.transition = "left 900ms ease, top 900ms ease, width 900ms ease, height 900ms ease";
          clone.style.left = `${targetLeft}px`;
          clone.style.top = `${targetTop}px`;
          clone.style.width = `${targetWidth}px`;
          clone.style.height = `${targetHeight}px`;
        });

        window.setTimeout(() => {
          metricsBagPage.scrollIntoView({ behavior: "auto", block: "start" });
          metricsBagPage.classList.add("is-active");
          overlay.innerHTML = "";
          methodIntroState.transitioningToBag = false;
        }, 930);
      };

      const goToGameInstant = () => {
        allowGameJump = true;
        gamePage.scrollIntoView({ behavior: "auto", block: "start" });
        window.setTimeout(() => {
          allowGameJump = false;
        }, 250);
      };

      playButton.addEventListener("click", goToGameInstant);
      playSkipButton.addEventListener("click", () => showMainInterface());

      function sectionVisibilityRatio(element) {
        const rect = element.getBoundingClientRect();
        const viewportH = window.innerHeight || document.documentElement.clientHeight;
        const visibleTop = Math.max(0, rect.top);
        const visibleBottom = Math.min(viewportH, rect.bottom);
        const visibleHeight = Math.max(0, visibleBottom - visibleTop);
        return rect.height > 0 ? visibleHeight / rect.height : 0;
      }

      function isPlayIntroActive() {
        return sectionVisibilityRatio(playIntroPage) > 0.6;
      }

      function isMethodIntroActive() {
        return sectionVisibilityRatio(methodIntroPage) > 0.6;
      }

      function isGameActive() {
        return sectionVisibilityRatio(gamePage) > 0.6;
      }

      function isMetricsBagActive() {
        return sectionVisibilityRatio(metricsBagPage) > 0.6;
      }

      function isInsightsActive() {
        return sectionVisibilityRatio(insightsPage) > 0.6;
      }

      window.addEventListener(
        "wheel",
        (event) => {
          if (allowGameJump) {
            return;
          }

          const nowMs = Date.now();
          if (nowMs - methodIntroState.lastWheelEventAtMs > 240) {
            methodIntroState.wheelGestureId += 1;
          }
          methodIntroState.lastWheelEventAtMs = nowMs;

          if (methodIntroState.transitioningToBag) {
            event.preventDefault();
            return;
          }

          if (isMethodIntroActive() && event.deltaY > 0 && methodIntroState.suppressMethodDownUntilIdle) {
            event.preventDefault();
            if (methodIntroState.methodDownIdleTimer) {
              window.clearTimeout(methodIntroState.methodDownIdleTimer);
            }
            methodIntroState.methodDownIdleTimer = window.setTimeout(() => {
              methodIntroState.suppressMethodDownUntilIdle = false;
              methodIntroState.methodDownIdleTimer = null;
            }, 180);
            return;
          }

          if (isMethodIntroActive() && !methodIntroState.revealed) {
            if (event.deltaY > 0) {
              event.preventDefault();
              revealMethodIntro();
            }
            return;
          }

          if (isMethodIntroActive() && methodIntroState.revealed && event.deltaY > 0) {
            if (Date.now() < methodIntroState.entryLockUntilMs) {
              event.preventDefault();
              return;
            }
            if (methodIntroState.waitingForSecondScrollToBag) {
              methodIntroState.waitingForSecondScrollToBag = false;
              event.preventDefault();
              return;
            }
            if (methodIntroState.wheelGestureId <= methodIntroState.revealGestureId) {
              event.preventDefault();
              return;
            }
            event.preventDefault();
            startMethodToBagTransition();
            return;
          }

          if (isMetricsBagActive() && event.deltaY < 0) {
            metricsBagPage.classList.remove("is-active");
            metricsBagPage.classList.remove("is-opened");
            hideMetricDetail(false);
          }

          if (isMetricsBagActive() && event.deltaY > 0 && !metricsBagState.canScrollToGameIntro) {
            event.preventDefault();
            return;
          }

          if (isPlayIntroActive()) {
            if (event.deltaY > 0) {
              event.preventDefault();
            }
            return;
          }

          if (isGameActive()) {
            if (event.deltaX > 40 && Math.abs(event.deltaX) > Math.abs(event.deltaY)) {
              playIntroPage.scrollIntoView({ behavior: "smooth", block: "start" });
              return;
            }
            if (event.deltaY < 0) {
              event.preventDefault();
            }
            return;
          }

          if (isInsightsActive() && Math.abs(event.deltaY) > 0) {
            event.preventDefault();
          }
        },
        { passive: false }
      );

      window.addEventListener(
        "keydown",
        (event) => {
          if (allowGameJump) {
            return;
          }
          if (methodIntroState.transitioningToBag) {
            event.preventDefault();
            return;
          }
          if (
            isMethodIntroActive()
            && !methodIntroState.revealed
            && ["ArrowDown", "PageDown", " ", "End"].includes(event.key)
          ) {
            event.preventDefault();
            revealMethodIntro();
            return;
          }
          if (
            isMethodIntroActive()
            && methodIntroState.revealed
            && ["ArrowDown", "PageDown", " ", "End"].includes(event.key)
          ) {
            if (Date.now() < methodIntroState.entryLockUntilMs) {
              event.preventDefault();
              return;
            }
            if (methodIntroState.waitingForSecondScrollToBag) {
              methodIntroState.waitingForSecondScrollToBag = false;
              event.preventDefault();
              return;
            }
            event.preventDefault();
            startMethodToBagTransition();
            return;
          }
          if (isPlayIntroActive() && ["ArrowDown", "PageDown", " ", "End"].includes(event.key)) {
            event.preventDefault();
            return;
          }
          if (isMetricsBagActive() && ["ArrowDown", "PageDown", " ", "End"].includes(event.key) && !metricsBagState.canScrollToGameIntro) {
            event.preventDefault();
            return;
          }
          if (isGameActive() && ["ArrowUp", "PageUp", "Home"].includes(event.key)) {
            event.preventDefault();
            return;
          }
          if (isInsightsActive() && ["ArrowUp", "ArrowDown", "PageUp", "PageDown", " ", "Home", "End"].includes(event.key)) {
            event.preventDefault();
          }
        },
        { passive: false }
      );
    }

    function initTitleCarousel() {
      const track = document.getElementById("boundary-carousel-track");
      if (!track) {
        return;
      }

      const modelFiles = Array.from({ length: 21 }, (_, index) => `./media/models/model${index + 1}.png`);
      const createImage = (filePath) => {
        const image = document.createElement("img");
        image.className = "title-model-image";
        image.src = filePath;
        image.alt = "";
        image.loading = "eager";
        image.decoding = "async";
        return image;
      };

      track.innerHTML = "";
      modelFiles.forEach((filePath) => track.appendChild(createImage(filePath)));
      modelFiles.forEach((filePath) => track.appendChild(createImage(filePath)));

      const updateShiftDistance = () => {
        const images = Array.from(track.querySelectorAll(".title-model-image"));
        const singleSet = images.slice(0, modelFiles.length);
        const totalWidth = singleSet.reduce((sum, image) => sum + image.getBoundingClientRect().width, 0);
        const totalGap = Math.max(0, singleSet.length - 1) * 3;
        track.style.setProperty("--title-carousel-shift", `${Math.round(totalWidth + totalGap)}px`);
      };

      const allImages = Array.from(track.querySelectorAll(".title-model-image"));
      let pendingLoads = allImages.length;
      const handleLoad = () => {
        pendingLoads -= 1;
        if (pendingLoads <= 0) {
          updateShiftDistance();
        }
      };

      allImages.forEach((image) => {
        if (image.complete) {
          handleLoad();
        } else {
          image.addEventListener("load", handleLoad, { once: true });
          image.addEventListener("error", handleLoad, { once: true });
        }
      });

      if (pendingLoads <= 0) {
        updateShiftDistance();
      }

      window.addEventListener("resize", updateShiftDistance);
    }


    async function loadData() {
      const [controlResponse, colorResponse] = await Promise.all([
        fetch(controlFile),
        fetch(colorFile)
      ]);

      if (!controlResponse.ok || !colorResponse.ok) {
        throw new Error("Could not load one or both Excel files from the Data folder.");
      }

      const [controlBuffer, colorBuffer] = await Promise.all([
        controlResponse.arrayBuffer(),
        colorResponse.arrayBuffer()
      ]);

      const controlRows = getFirstSheetRows(readWorkbookFromBuffer(controlBuffer));
      const colorRows = getFirstSheetRows(readWorkbookFromBuffer(colorBuffer));

      const controlData = parseControlRows(controlRows);
      const colorMap = parseColorRows(colorRows);

      return controlData.map((item) => ({
        ...item,
        color: colorMap.get(normalizeBrand(item.brand)) || "#000000"
      }));
    }

    let currentPoints = [];

    async function init() {
      try {
        currentPoints = await loadData();
        renderIntroChart();
        createIntroTokens();
        renderLogoScatter(currentPoints);

        initPlayGameNavigation();
        initTitleCarousel();
        initScrollTransitions();
      } catch (error) {
        console.error(error);
      }
    }

    window.addEventListener("resize", () => {
      const introView = document.getElementById("intro-view");
      if (introView && introView.style.display !== "none") {
        renderIntroChart();
        if (introState.revealed) {
          revealIntroCorrectPlacements(true);
        } else {
          placeIntroTokensInTray();
        }
      }

      if (currentPoints.length > 0) {
        renderLogoScatter(currentPoints);
      }
    });

    init();
  