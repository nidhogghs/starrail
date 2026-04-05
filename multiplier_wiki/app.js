(function () {
  const data = window.MULTIPLIER_WIKI_DATA;
  if (!data || !Array.isArray(data.roles)) {
    document.body.innerHTML = "<div class='page-shell'><p class='empty-state'>未找到 data.js，请先运行生成脚本。</p></div>";
    return;
  }

  const bucketLabels = {
    "基础倍率_攻击%": "攻击倍率%",
    "基础倍率_生命%": "生命倍率%",
    "基础倍率_防御%": "防御倍率%",
    "基础倍率_欢愉伤害%": "欢愉伤害倍率%",
    "基础倍率_原伤害%": "原伤害倍率%",
    "基础倍率_真实伤害%": "真实伤害倍率%",
    "基础倍率_击破特攻系数": "击破特攻系数",
    "基础倍率_欢愉度系数": "欢愉度系数",
    "基础倍率_其他%": "其他基础倍率%",
    "攻击%": "攻击%",
    "生命%": "生命%",
    "防御%": "防御%",
    "速度%": "速度%",
    "速度点": "速度点",
    "增伤区%": "增伤区%",
    "易伤区%": "易伤区%",
    "暴击率%": "暴击率%",
    "暴击伤害%": "暴击伤害%",
    "防御穿透/减防%": "减防/无视防御%",
    "抗穿/减抗%": "抗穿/减抗%",
    "击破特攻%": "击破特攻%",
    "击破伤害提高%": "击破伤害提高%",
    "超击破伤害提高%": "超击破伤害提高%",
    "弱点击破效率/无视弱点削韧%": "弱点击破效率%",
    "欢愉度%": "欢愉度%",
    "笑点": "笑点",
    "增笑%": "增笑%",
    "伤害减免/受击减伤%": "伤害减免%",
  };

  const rolesByName = Object.fromEntries(data.roles.map((role) => [role.name, role]));
  const roleNames = data.roles.map((role) => role.name);
  const defaults = data.defaults;
  const summaryColumns = data.summaryColumns;

  const state = {
    browseRole: roleNames[0],
    browseRowId: null,
    team: [null, null, null, null],
    focusRole: "",
    focusRowId: "",
    settings: {
      attackerLevel: defaults.attackerLevel,
      enemyLevel: defaults.enemyLevel,
      enemyResistance: defaults.enemyResistance,
      enemyDamageReduction: defaults.enemyDamageReduction,
      enemyBroken: defaults.enemyBroken,
      actionValueLimit: defaults.actionValueLimit,
      baseCritRate: defaults.baseCritRate,
      baseCritDamage: defaults.baseCritDamage,
    },
    specialSettings: {},
  };

  const els = {
    heroStats: document.getElementById("hero-stats"),
    browseRole: document.getElementById("browse-role"),
    roleMeta: document.getElementById("role-meta"),
    rowList: document.getElementById("row-list"),
    rowDetail: document.getElementById("row-detail"),
    teamSlots: document.getElementById("team-slots"),
    focusRole: document.getElementById("focus-role"),
    focusRow: document.getElementById("focus-row"),
    attackerLevel: document.getElementById("attacker-level"),
    enemyLevel: document.getElementById("enemy-level"),
    actionValueLimit: document.getElementById("action-value-limit"),
    enemyResistance: document.getElementById("enemy-resistance"),
    enemyDamageReduction: document.getElementById("enemy-damage-reduction"),
    enemyBroken: document.getElementById("enemy-broken"),
    baseCritRate: document.getElementById("base-crit-rate"),
    baseCritDamage: document.getElementById("base-crit-damage"),
    specialSettings: document.getElementById("special-settings"),
    resultCards: document.getElementById("result-cards"),
    formulaBreakdown: document.getElementById("formula-breakdown"),
    timelinePanel: document.getElementById("timeline-panel"),
    bucketTableBody: document.getElementById("bucket-table-body"),
    appliedBuffs: document.getElementById("applied-buffs"),
  };

  function isFiniteNumber(value) {
    return Number.isFinite(Number(value));
  }

  function number(value, digits = 2) {
    const numeric = Number(value || 0);
    if (Number.isInteger(numeric)) {
      return String(numeric);
    }
    return numeric.toFixed(digits).replace(/\.?0+$/, "");
  }

  function cloneBucket(bucket) {
    const next = {};
    summaryColumns.forEach((column) => {
      next[column] = Number(bucket?.[column] || 0);
    });
    return next;
  }

  function addBuckets(target, source) {
    summaryColumns.forEach((column) => {
      target[column] = Number(target[column] || 0) + Number(source?.[column] || 0);
    });
    return target;
  }

  function nonZeroBuckets(bucket) {
    return summaryColumns
      .filter((column) => Number(bucket[column] || 0) !== 0)
      .map((column) => [column, Number(bucket[column])]);
  }

  function findRole(roleName) {
    return rolesByName[roleName] || null;
  }

  function selectedTeamRoles() {
    return state.team.filter(Boolean).filter((roleName, index, self) => self.indexOf(roleName) === index);
  }

  function damageRows(role) {
    return role.rows.filter((row) => row.hasBaseDamage);
  }

  function supportRows(role) {
    return role.rows.filter((row) => row.rowType === "增益");
  }

  function getRowById(role, rowId) {
    return role.rows.find((row) => row.id === rowId) || null;
  }

  function roleElement(roleName) {
    return findRole(roleName)?.element || "";
  }

  function isQuantumTeam(teamRoles) {
    return teamRoles.filter((roleName) => roleElement(roleName) === "量子").length >= 3;
  }

  function supportApplies(providerRole, row, focusRoleName, teamRoles) {
    const scope = row.targetScope || "";
    if (!scope) {
      return true;
    }
    if (scope.includes("我方全体")) {
      return true;
    }
    if (scope.includes("指定我方单体") || scope.includes("我方单体") || scope.includes("技能目标")) {
      return true;
    }
    if (scope.includes("持有【谜诡】")) {
      return true;
    }
    if (scope.includes("除自身外队友")) {
      return providerRole.name !== focusRoleName;
    }
    if (scope.includes("下一个行动的我方其他目标")) {
      return providerRole.name !== focusRoleName;
    }
    if (scope.includes("量子角色")) {
      return roleElement(focusRoleName) === "量子" && isQuantumTeam(teamRoles);
    }
    return true;
  }

  function calculateDefenseMultiplier(attackerLevel, enemyLevel, defenseIgnore) {
    const attacker = Number(attackerLevel || 80);
    const enemy = Number(enemyLevel || 80);
    const defense = Math.max(0, (200 + 10 * enemy) * (1 - Math.min(100, Math.max(0, defenseIgnore)) / 100));
    return 1 - defense / (defense + 200 + 10 * attacker);
  }

  function calculateResistanceMultiplier(enemyResistance, resistancePenetration) {
    const finalResistance = Math.max(-100, Math.min(90, Number(enemyResistance || 0) - Number(resistancePenetration || 0)));
    return 1 - finalResistance / 100;
  }

  function calculateDamageReductionMultiplier(enemyDamageReduction, enemyBroken) {
    return (1 - Number(enemyDamageReduction || 0) / 100) * (enemyBroken ? 1 : 0.9);
  }

  function calculateCritMultiplier(baseCritRate, baseCritDamage, bonusCritRate, bonusCritDamage) {
    const critRate = Math.max(0, Math.min(1, (Number(baseCritRate || 0) + Number(bonusCritRate || 0)) / 100));
    const critDamage = Math.max(0, (Number(baseCritDamage || 0) + Number(bonusCritDamage || 0)) / 100);
    return {
      critRate,
      critDamage,
      expected: 1 + critRate * critDamage,
    };
  }

  function calculateBaseIndex(bucket) {
    const attackPart = (bucket["基础倍率_攻击%"] / 100) * (1 + bucket["攻击%"] / 100);
    const hpPart = (bucket["基础倍率_生命%"] / 100) * (1 + bucket["生命%"] / 100);
    const defensePart = (bucket["基础倍率_防御%"] / 100) * (1 + bucket["防御%"] / 100);
    const breakPart = (bucket["基础倍率_击破特攻系数"] / 100) * (1 + bucket["击破特攻%"] / 100);
    const elationPart = (bucket["基础倍率_欢愉度系数"] / 100) * (1 + bucket["欢愉度%"] / 100);
    const otherPart =
      bucket["基础倍率_欢愉伤害%"] / 100 +
      bucket["基础倍率_原伤害%"] / 100 +
      bucket["基础倍率_真实伤害%"] / 100 +
      bucket["基础倍率_其他%"] / 100;
    const total = attackPart + hpPart + defensePart + breakPart + elationPart + otherPart;
    return {
      total: total || 0,
      attackPart,
      hpPart,
      defensePart,
      breakPart,
      elationPart,
      otherPart,
    };
  }

  function calculateRowResult(focusRole, row, teamRoles) {
    const totalBucket = cloneBucket(row.totalBucket);
    const appliedBuffs = [];

    teamRoles.forEach((roleName) => {
      const role = findRole(roleName);
      if (!role) {
        return;
      }
      supportRows(role).forEach((supportRow) => {
        if (!supportApplies(role, supportRow, focusRole.name, teamRoles)) {
          return;
        }
        addBuckets(totalBucket, supportRow.totalBucket);
        appliedBuffs.push({ provider: role.name, row: supportRow });
      });
    });

    const baseIndex = calculateBaseIndex(totalBucket);
    const damageBonusMultiplier = 1 + totalBucket["增伤区%"] / 100;
    const vulnerabilityMultiplier = 1 + totalBucket["易伤区%"] / 100;
    const crit = calculateCritMultiplier(
      state.settings.baseCritRate,
      state.settings.baseCritDamage,
      totalBucket["暴击率%"],
      totalBucket["暴击伤害%"]
    );
    const defenseMultiplier = calculateDefenseMultiplier(
      state.settings.attackerLevel,
      state.settings.enemyLevel,
      totalBucket["防御穿透/减防%"]
    );
    const resistanceMultiplier = calculateResistanceMultiplier(
      state.settings.enemyResistance,
      totalBucket["抗穿/减抗%"]
    );
    const damageReductionMultiplier = calculateDamageReductionMultiplier(
      state.settings.enemyDamageReduction,
      state.settings.enemyBroken
    );

    const finalIndex =
      baseIndex.total *
      damageBonusMultiplier *
      vulnerabilityMultiplier *
      crit.expected *
      defenseMultiplier *
      resistanceMultiplier *
      damageReductionMultiplier;

    return {
      totalBucket,
      appliedBuffs,
      baseIndex,
      damageBonusMultiplier,
      vulnerabilityMultiplier,
      crit,
      defenseMultiplier,
      resistanceMultiplier,
      damageReductionMultiplier,
      finalIndex,
    };
  }

  function calculateEffectiveSpeed(roleName, teamRoles) {
    const role = findRole(roleName);
    if (!role) {
      return 100;
    }
    let speedPct = 0;
    let speedFlat = 0;
    teamRoles.forEach((providerName) => {
      const providerRole = findRole(providerName);
      if (!providerRole) {
        return;
      }
      supportRows(providerRole).forEach((row) => {
        if (!supportApplies(providerRole, row, roleName, teamRoles)) {
          return;
        }
        speedPct += Number(row.totalBucket["速度%"] || 0);
        speedFlat += Number(row.totalBucket["速度点"] || 0);
      });
    });
    return role.baseSpeed * (1 + speedPct / 100) + speedFlat;
  }

  function calculateResult(focusRole, focusRow) {
    if (!focusRole || !focusRow) {
      return null;
    }

    const teamRoles = selectedTeamRoles();
    const rowResult = calculateRowResult(focusRole, focusRow, teamRoles);
    const actionSimulation = window.ROLE_ACTION_LOGIC
      ? window.ROLE_ACTION_LOGIC.simulate({
          focusRole,
          focusRow,
          teamRoles,
          rolesByName,
          settings: state.settings,
          specialSettings: state.specialSettings,
          getDamageIndex(targetRow) {
            return calculateRowResult(focusRole, targetRow, teamRoles).finalIndex;
          },
          getEffectiveSpeed(roleName) {
            return calculateEffectiveSpeed(roleName, teamRoles);
          },
          formatNumber: number,
        })
      : null;

    return {
      ...rowResult,
      actionSimulation,
    };
  }

  function createOption(value, label) {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = label;
    return option;
  }

  function renderHeroStats() {
    const stats = [
      { label: "角色数", value: data.roles.length, note: "当前项目内已收录的角色" },
      { label: "技能/条目数", value: data.roles.reduce((sum, role) => sum + role.rows.length, 0), note: "角色条目与辅助增益合计" },
      { label: "默认敌抗", value: `${defaults.enemyResistance}%`, note: "可在计算器里随时改" },
      { label: "数据生成时间", value: data.generatedAt.replace("T", " "), note: "和 Excel 输出保持同一批数据" },
    ];

    els.heroStats.innerHTML = "";
    stats.forEach((item) => {
      const card = document.createElement("div");
      card.className = "hero-stat-card";
      card.innerHTML = `<span>${item.label}</span><strong>${item.value}</strong><span>${item.note}</span>`;
      els.heroStats.appendChild(card);
    });
  }

  function renderBrowseRoleSelect() {
    els.browseRole.innerHTML = "";
    roleNames.forEach((roleName) => {
      els.browseRole.appendChild(createOption(roleName, roleName));
    });
    els.browseRole.value = state.browseRole;
  }

  function renderRoleMeta(role) {
    els.roleMeta.innerHTML = "";
    [
      role.path ? `命途：${role.path}` : "",
      role.element ? `属性：${role.element}` : "",
      role.baseSpeed ? `基础速度：${role.baseSpeed}` : "",
      role.isSupport ? "定位：辅助" : "定位：输出/混合",
      `条目数：${role.rows.length}`,
    ]
      .filter(Boolean)
      .forEach((text) => {
        const pill = document.createElement("div");
        pill.className = "meta-pill";
        pill.textContent = text;
        els.roleMeta.appendChild(pill);
      });
  }

  function renderRowList(role) {
    const rows = role.rows;
    if (!state.browseRowId && rows.length) {
      state.browseRowId = rows[0].id;
    }

    els.rowList.innerHTML = "";
    rows.forEach((row) => {
      const card = document.createElement("button");
      card.type = "button";
      card.className = `row-card${row.id === state.browseRowId ? " active" : ""}`;
      const totalEntries = nonZeroBuckets(row.totalBucket).slice(0, 4);
      card.innerHTML = `
        <div class="row-card-head">
          <h3>${row.sourceName}</h3>
          <span class="row-type">${row.rowType}</span>
        </div>
        <div class="row-tags">
          ${totalEntries.map(([key, value]) => `<span class="tag">${bucketLabels[key]} ${number(value)}</span>`).join("")}
        </div>
        <p class="row-snippet">${(row.baseText || "").slice(0, 120)}${(row.baseText || "").length > 120 ? "..." : ""}</p>
      `;
      card.addEventListener("click", () => {
        state.browseRowId = row.id;
        renderBrowser();
      });
      els.rowList.appendChild(card);
    });
  }

  function renderRowDetail(role) {
    const row = getRowById(role, state.browseRowId) || role.rows[0];
    if (!row) {
      els.rowDetail.innerHTML = "<div class='empty-state'>这个角色当前没有条目。</div>";
      return;
    }
    state.browseRowId = row.id;

    const totalBuckets = nonZeroBuckets(row.totalBucket)
      .map(([key, value]) => `<div class="bucket-item"><span>${bucketLabels[key]}</span><strong>${number(value)}</strong></div>`)
      .join("");

    const notes = row.notes && row.notes.length
      ? `<div class="bucket-list">${row.notes.map((note) => `<div class="bucket-item"><span>备注</span><strong>${note}</strong></div>`).join("")}</div>`
      : "";

    const sourceLink = row.sourceUrl
      ? `<a class="link-button" href="${row.sourceUrl}" target="_blank" rel="noreferrer">查看来源</a>`
      : "";

    els.rowDetail.innerHTML = `
      <div class="detail-grid">
        <div class="detail-card">
          <span>来源</span>
          <strong>${row.sourceCol} / ${row.sourceName}</strong>
        </div>
        <div class="detail-card">
          <span>作用对象 / 持续</span>
          <strong>${row.targetScope || "自身/未单列"} · ${row.duration || "即时/常驻未单列"}</strong>
        </div>
      </div>
      <div class="bucket-list" style="margin-top:16px;">
        <div class="bucket-item"><span>技能描述</span><strong>${row.baseText || "无"}</strong></div>
      </div>
      <div class="bucket-list" style="margin-top:14px;">${totalBuckets || "<div class='empty-state'>这个条目没有可量化乘区。</div>"}</div>
      ${notes}
      <div style="margin-top:14px;">${sourceLink}</div>
    `;
  }

  function renderBrowser() {
    const role = findRole(state.browseRole);
    if (!role) {
      return;
    }
    renderRoleMeta(role);
    renderRowList(role);
    renderRowDetail(role);
  }

  function renderTeamSlots() {
    els.teamSlots.innerHTML = "";
    state.team.forEach((roleName, index) => {
      const wrapper = document.createElement("div");
      wrapper.className = "slot-card";
      const select = document.createElement("select");
      select.appendChild(createOption("", "留空"));
      roleNames.forEach((name) => {
        select.appendChild(createOption(name, name));
      });
      select.value = roleName || "";
      select.addEventListener("change", (event) => {
        state.team[index] = event.target.value || null;
        syncFocusRole();
        renderTeamSlots();
        renderCalculation();
      });
      wrapper.innerHTML = `<p class="slot-label">阵容位 ${index + 1}</p>`;
      wrapper.appendChild(select);
      els.teamSlots.appendChild(wrapper);
    });
  }

  function syncFocusRole() {
    const teamRoles = selectedTeamRoles();
    if (!teamRoles.length) {
      state.focusRole = "";
      state.focusRowId = "";
      return;
    }
    if (!teamRoles.includes(state.focusRole)) {
      state.focusRole = teamRoles[0];
      state.focusRowId = "";
    }
  }

  function renderFocusSelectors() {
    syncFocusRole();
    const teamRoles = selectedTeamRoles();
    els.focusRole.innerHTML = "";
    els.focusRole.appendChild(createOption("", "先选择阵容"));
    teamRoles.forEach((roleName) => {
      els.focusRole.appendChild(createOption(roleName, roleName));
    });
    els.focusRole.value = state.focusRole;

    els.focusRow.innerHTML = "";
    if (!state.focusRole) {
      els.focusRow.appendChild(createOption("", "先选择主C"));
      return;
    }

    const role = findRole(state.focusRole);
    const rows = damageRows(role);
    els.focusRow.appendChild(createOption("", "选择技能"));
    rows.forEach((row) => {
      els.focusRow.appendChild(createOption(row.id, `${row.rowType} · ${row.sourceName}`));
    });

    if (!rows.find((row) => row.id === state.focusRowId)) {
      state.focusRowId = rows[0]?.id || "";
    }
    els.focusRow.value = state.focusRowId;
  }

  function ensureSpecialSettingDefaults(roleName) {
    const schema = window.ROLE_ACTION_LOGIC ? window.ROLE_ACTION_LOGIC.getSchema(roleName) : [];
    state.specialSettings[roleName] = state.specialSettings[roleName] || {};
    schema.forEach((item) => {
      if (state.specialSettings[roleName][item.key] === undefined) {
        state.specialSettings[roleName][item.key] = item.defaultValue;
      }
    });
  }

  function renderSpecialSettings(focusRole) {
    els.specialSettings.innerHTML = "";
    if (!focusRole || !window.ROLE_ACTION_LOGIC) {
      return;
    }

    const schema = window.ROLE_ACTION_LOGIC.getSchema(focusRole.name);
    if (!schema.length) {
      els.specialSettings.innerHTML = "<div class='empty-state'>这个角色当前没有额外行动轴参数，按默认逻辑模拟。</div>";
      return;
    }

    ensureSpecialSettingDefaults(focusRole.name);
    const title = document.createElement("div");
    title.className = "panel-kicker";
    title.textContent = "角色专项行动逻辑";
    els.specialSettings.appendChild(title);

    schema.forEach((item) => {
      const wrapper = document.createElement("label");
      wrapper.className = "field";
      wrapper.innerHTML = `<span>${item.label}</span>`;
      const currentValue = state.specialSettings[focusRole.name][item.key];
      let input;
      if (item.type === "checkbox") {
        input = document.createElement("input");
        input.type = "checkbox";
        input.checked = Boolean(currentValue);
        input.addEventListener("change", (event) => {
          state.specialSettings[focusRole.name][item.key] = event.target.checked;
          renderCalculation();
        });
      } else {
        input = document.createElement("input");
        input.type = "number";
        input.min = item.min;
        input.max = item.max;
        input.step = item.step || 1;
        input.value = currentValue;
        input.addEventListener("input", (event) => {
          if (!isFiniteNumber(event.target.value)) {
            return;
          }
          state.specialSettings[focusRole.name][item.key] = Number(event.target.value);
          renderCalculation();
        });
      }
      wrapper.appendChild(input);
      if (item.help) {
        const hint = document.createElement("span");
        hint.textContent = item.help;
        wrapper.appendChild(hint);
      }
      els.specialSettings.appendChild(wrapper);
    });
  }

  function renderSettings() {
    els.attackerLevel.value = state.settings.attackerLevel;
    els.enemyLevel.value = state.settings.enemyLevel;
    els.actionValueLimit.value = state.settings.actionValueLimit;
    els.enemyResistance.value = state.settings.enemyResistance;
    els.enemyDamageReduction.value = state.settings.enemyDamageReduction;
    els.enemyBroken.checked = state.settings.enemyBroken;
    els.baseCritRate.value = state.settings.baseCritRate;
    els.baseCritDamage.value = state.settings.baseCritDamage;
  }

  function renderResult(result, focusRole, focusRow) {
    if (!result || !focusRole || !focusRow) {
      els.resultCards.innerHTML = "<div class='empty-state'>先在阵容里选主C和技能，结果区才会开始计算。</div>";
      els.formulaBreakdown.innerHTML = "";
      els.timelinePanel.innerHTML = "";
      els.bucketTableBody.innerHTML = "";
      els.appliedBuffs.innerHTML = "";
      return;
    }

    const totalDamageOnAxis = result.actionSimulation ? result.actionSimulation.totalDamage : 0;
    const cards = [
      ["最终相对输出指数", result.finalIndex],
      ["行动轴总伤", totalDamageOnAxis],
      ["行动值长度", state.settings.actionValueLimit],
      ["基础倍率指数", result.baseIndex.total],
      ["增伤乘区", result.damageBonusMultiplier],
      ["防御乘区", result.defenseMultiplier],
      ["抗性乘区", result.resistanceMultiplier],
      ["暴击期望乘区", result.crit.expected],
    ];

    els.resultCards.innerHTML = cards
      .map(
        ([label, value]) =>
          `<div class="result-card"><span>${label}</span><strong>${number(value, 4)}</strong></div>`
      )
      .join("");

    const lines = [
      ["主C", `${focusRole.name} · ${focusRow.sourceName}`],
      ["基础倍率指数", number(result.baseIndex.total, 4)],
      ["增伤乘区", number(result.damageBonusMultiplier, 4)],
      ["易伤乘区", number(result.vulnerabilityMultiplier, 4)],
      ["暴击期望乘区", `${number(result.crit.expected, 4)}（暴击率 ${number(result.crit.critRate * 100)}% / 暴伤 ${number(result.crit.critDamage * 100)}%）`],
      ["防御乘区", number(result.defenseMultiplier, 4)],
      ["抗性乘区", number(result.resistanceMultiplier, 4)],
      ["减伤乘区", number(result.damageReductionMultiplier, 4)],
    ];
    els.formulaBreakdown.innerHTML = lines
      .map(([label, value]) => `<div class="formula-line"><span>${label}</span><strong>${value}</strong></div>`)
      .join("");

    if (result.actionSimulation) {
      const actionSummary = Object.entries(result.actionSimulation.actionCounts || {})
        .map(([label, count]) => `<div class="formula-line"><span>${label} 行动次数</span><strong>${count}</strong></div>`)
        .join("");
      const timelineEvents = (result.actionSimulation.timeline || [])
        .slice(0, 24)
        .map(
          (item) =>
            `<div class="timeline-event"><span>AV ${number(item.av, 2)} · ${item.label}</span><strong>${item.detail || ""}</strong></div>`
        )
        .join("");
      const notes = (result.actionSimulation.notes || [])
        .map((note) => `<div class="timeline-event"><span>说明</span><strong>${note}</strong></div>`)
        .join("");
      els.timelinePanel.innerHTML = `
        <div class="timeline-card">
          <h3>行动轴</h3>
          <div class="formula-breakdown">${actionSummary || "<div class='empty-state'>当前没有可统计的行动次数。</div>"}</div>
        </div>
        <div class="timeline-card">
          <h3>行动事件</h3>
          <div class="timeline-events">${timelineEvents || "<div class='empty-state'>当前行动值长度内没有触发行。</div>"}${notes}</div>
        </div>
      `;
    } else {
      els.timelinePanel.innerHTML = "";
    }

    els.bucketTableBody.innerHTML = nonZeroBuckets(result.totalBucket)
      .sort((a, b) => Math.abs(b[1]) - Math.abs(a[1]))
      .map(([key, value]) => `<tr><td>${bucketLabels[key]}</td><td>${number(value)}</td></tr>`)
      .join("");

    if (!result.appliedBuffs.length) {
      els.appliedBuffs.innerHTML = "<div class='empty-state'>当前阵容没有额外套到主C身上的辅助增益。</div>";
      return;
    }

    els.appliedBuffs.innerHTML = result.appliedBuffs
      .map(({ provider, row }) => {
        const tags = [row.targetScope || "目标未单列", row.duration || "持续未单列"]
          .map((text) => `<span class="buff-tag">${text}</span>`)
          .join("");
        const bucketSummary = nonZeroBuckets(row.totalBucket)
          .slice(0, 5)
          .map(([key, value]) => `<span class="buff-tag">${bucketLabels[key]} ${number(value)}</span>`)
          .join("");
        return `
          <div class="buff-card">
            <h4>${provider} · ${row.sourceName}</h4>
            <div class="buff-tags">${tags}${bucketSummary}</div>
            <p>${row.baseText || "无描述"}</p>
          </div>
        `;
      })
      .join("");
  }

  function renderCalculation() {
    renderFocusSelectors();
    renderSettings();

    const focusRole = findRole(state.focusRole);
    const focusRow = focusRole ? getRowById(focusRole, state.focusRowId) || damageRows(focusRole)[0] : null;
    if (focusRow) {
      state.focusRowId = focusRow.id;
    }
    renderSpecialSettings(focusRole);
    const result = calculateResult(focusRole, focusRow);
    renderResult(result, focusRole, focusRow);
  }

  function attachEvents() {
    els.browseRole.addEventListener("change", (event) => {
      state.browseRole = event.target.value;
      state.browseRowId = "";
      renderBrowser();
    });

    els.focusRole.addEventListener("change", (event) => {
      state.focusRole = event.target.value;
      state.focusRowId = "";
      renderCalculation();
    });

    els.focusRow.addEventListener("change", (event) => {
      state.focusRowId = event.target.value;
      renderCalculation();
    });

    [
      ["attackerLevel", els.attackerLevel],
      ["enemyLevel", els.enemyLevel],
      ["actionValueLimit", els.actionValueLimit],
      ["enemyResistance", els.enemyResistance],
      ["enemyDamageReduction", els.enemyDamageReduction],
      ["baseCritRate", els.baseCritRate],
      ["baseCritDamage", els.baseCritDamage],
    ].forEach(([key, input]) => {
      input.addEventListener("input", (event) => {
        if (!isFiniteNumber(event.target.value)) {
          return;
        }
        state.settings[key] = Number(event.target.value);
        renderCalculation();
      });
    });

    els.enemyBroken.addEventListener("change", (event) => {
      state.settings.enemyBroken = event.target.checked;
      renderCalculation();
    });
  }

  function init() {
    renderHeroStats();
    renderBrowseRoleSelect();
    renderTeamSlots();
    renderBrowser();
    renderCalculation();
    attachEvents();
  }

  init();
})();
