(function () {
  function baseAV(speed) {
    return 10000 / Math.max(1, speed);
  }

  function clamp(value, min, max) {
    return Math.max(min, Math.min(max, value));
  }

  function findRow(role, predicate) {
    return role?.rows?.find(predicate) || null;
  }

  function defaultSchema() {
    return [];
  }

  function createBattle(context) {
    const battle = {
      context,
      elapsedAV: 0,
      actors: [],
      pendingImmediate: [],
      timeline: [],
      damageEvents: [],
      actionCounts: {},
      notes: [],
      focusDamage: 0,
      teamOrder: context.teamRoles.slice(),
    };

    battle.addActor = function addActor(actor) {
      actor.baseSpeed = actor.baseSpeed || 100;
      actor.currentSpeed = actor.currentSpeed || actor.baseSpeed;
      actor.currentAV = actor.currentAV ?? baseAV(actor.currentSpeed);
      actor.active = actor.active !== false;
      actor.turns = 0;
      actor.tags = actor.tags || [];
      actor.state = actor.state || {};
      battle.actors.push(actor);
      return actor;
    };

    battle.resetActorAV = function resetActorAV(actor) {
      actor.currentAV = baseAV(actor.currentSpeed);
    };

    battle.setActorSpeed = function setActorSpeed(actor, newSpeed) {
      const safeNewSpeed = Math.max(1, newSpeed);
      const oldSpeed = Math.max(1, actor.currentSpeed);
      actor.currentAV = actor.currentAV * (oldSpeed / safeNewSpeed);
      actor.currentSpeed = safeNewSpeed;
    };

    battle.advanceActor = function advanceActor(actor, percent) {
      actor.currentAV = Math.max(0, actor.currentAV - baseAV(actor.currentSpeed) * (percent / 100));
    };

    battle.immediateActor = function immediateActor(actor, reason) {
      if (!actor || !actor.active) {
        return;
      }
      battle.pendingImmediate.push({ actor, reason: reason || "立即行动" });
      actor.currentAV = 0;
    };

    battle.addTimeline = function addTimeline(label, detail) {
      battle.timeline.push({
        av: Number(battle.elapsedAV.toFixed(2)),
        label,
        detail,
      });
    };

    battle.addDamage = function addDamage(actor, row, multiplier, detail) {
      const perActionDamage = context.getDamageIndex(row);
      const totalDamage = perActionDamage * (multiplier || 1);
      battle.damageEvents.push({
        av: Number(battle.elapsedAV.toFixed(2)),
        actor: actor.label,
        source: row ? `${row.rowType}·${row.sourceName}` : actor.label,
        damage: totalDamage,
        detail,
      });
      if (actor.isFocusOwned) {
        battle.focusDamage += totalDamage;
      }
      return totalDamage;
    };

    battle.recordAction = function recordAction(actor) {
      actor.turns += 1;
      battle.actionCounts[actor.label] = (battle.actionCounts[actor.label] || 0) + 1;
    };

    return battle;
  }

  function runBattle(battle) {
    const actionLimit = battle.context.settings.actionValueLimit;
    while (true) {
      const immediate = battle.pendingImmediate.shift();
      if (immediate) {
        executeActorTurn(battle, immediate.actor, true, immediate.reason);
        continue;
      }

      const activeActors = battle.actors.filter((actor) => actor.active);
      if (!activeActors.length) {
        break;
      }

      let nextActor = activeActors[0];
      for (const actor of activeActors) {
        if (actor.currentAV < nextActor.currentAV) {
          nextActor = actor;
        }
      }

      if (battle.elapsedAV + nextActor.currentAV > actionLimit) {
        break;
      }

      const delta = nextActor.currentAV;
      battle.elapsedAV += delta;
      activeActors.forEach((actor) => {
        actor.currentAV -= delta;
      });
      executeActorTurn(battle, nextActor, false, "");
    }

    return {
      totalDamage: battle.focusDamage,
      elapsedAV: Number(battle.elapsedAV.toFixed(2)),
      timeline: battle.timeline,
      damageEvents: battle.damageEvents,
      actionCounts: battle.actionCounts,
      notes: battle.notes,
    };
  }

  function executeActorTurn(battle, actor, immediate, reason) {
    if (!actor.active) {
      return;
    }

    battle.recordAction(actor);
    const result = actor.onTurn
      ? actor.onTurn({ battle, actor, immediate, reason, context: battle.context }) || {}
      : {};

    if (actor.active && result.resetAV !== false) {
      battle.resetActorAV(actor);
    }
  }

  function standardFocusActor(context, roleName, row) {
    const role = context.rolesByName[roleName];
    return {
      id: `${roleName}-main`,
      label: roleName,
      roleName,
      baseSpeed: role.baseSpeed,
      currentSpeed: context.getEffectiveSpeed(roleName),
      isFocusOwned: true,
      state: {},
      onTurn({ battle, actor }) {
        battle.addDamage(actor, row, 1, row.sourceName);
        battle.addTimeline(actor.label, `${row.sourceName} 造成 ${context.formatNumber(context.getDamageIndex(row), 4)} 指数伤害`);
        return {};
      },
    };
  }

  function supportActor(context, roleName, onTurn) {
    const role = context.rolesByName[roleName];
    return {
      id: `${roleName}-support`,
      label: roleName,
      roleName,
      baseSpeed: role.baseSpeed,
      currentSpeed: context.getEffectiveSpeed(roleName),
      isFocusOwned: false,
      state: {},
      onTurn,
    };
  }

  function getSpecialSettings(settings, roleName) {
    return settings?.[roleName] || {};
  }

  function schemaSeele() {
    return [
      {
        key: "resurgenceRate",
        label: "再现触发率%",
        type: "number",
        min: 0,
        max: 100,
        step: 5,
        defaultValue: 0,
        help: "按每个常规回合触发一次再现的期望概率处理，额外回合不会继续连锁触发。",
      },
    ];
  }

  function simulateSeele(context) {
    const battle = createBattle(context);
    const settings = getSpecialSettings(context.specialSettings, "希儿");
    const rate = clamp(Number(settings.resurgenceRate || 0), 0, 100);
    const actor = standardFocusActor(context, "希儿", context.focusRow);
    actor.state.resurgenceMeter = 0;
    actor.onTurn = function onTurn({ battle: currentBattle, actor: self, immediate }) {
      currentBattle.addDamage(self, context.focusRow, 1, context.focusRow.sourceName);
      currentBattle.addTimeline(self.label, `${context.focusRow.sourceName}${immediate ? "（再现额外回合）" : ""}`);
      if (!immediate && rate > 0) {
        self.state.resurgenceMeter += rate;
        if (self.state.resurgenceMeter >= 100) {
          self.state.resurgenceMeter -= 100;
          currentBattle.immediateActor(self, "再现");
        }
      }
      return {};
    };
    battle.addActor(actor);
    addSupportActors(battle, context);
    return runBattle(battle);
  }

  function schemaJingYuan() {
    return [
      {
        key: "includeLightningLord",
        label: "计入神君",
        type: "checkbox",
        defaultValue: true,
        help: "景元行动外，额外模拟神君的独立行动。",
      },
    ];
  }

  function simulateJingYuan(context) {
    const battle = createBattle(context);
    const settings = getSpecialSettings(context.specialSettings, "景元");
    const includeLightningLord = settings.includeLightningLord !== false;
    const mainActor = standardFocusActor(context, "景元", context.focusRow);
    const lightningLordRow = findRow(context.focusRole, (row) => row.sourceName === "斩勘神形");
    let lightningLord = null;

    if (includeLightningLord && lightningLordRow) {
      lightningLord = battle.addActor({
        id: "景元-神君",
        label: "神君",
        roleName: "景元",
        baseSpeed: 60,
        currentSpeed: 60,
        isFocusOwned: true,
        state: { hits: 3 },
        onTurn({ battle: currentBattle, actor }) {
          const hits = actor.state.hits;
          currentBattle.addDamage(actor, lightningLordRow, hits, `神君 ${hits} 段`);
          currentBattle.addTimeline(actor.label, `神君行动，段数 ${hits}`);
          actor.state.hits = 3;
          currentBattle.setActorSpeed(actor, 60);
          return {};
        },
      });
    }

    mainActor.onTurn = function onTurn({ battle: currentBattle, actor }) {
      currentBattle.addDamage(actor, context.focusRow, 1, context.focusRow.sourceName);
      currentBattle.addTimeline(actor.label, `${context.focusRow.sourceName}`);
      if (lightningLord) {
        const extraHits = context.focusRow.rowType === "战技" ? 2 : context.focusRow.rowType === "终结技" ? 3 : 0;
        if (extraHits > 0) {
          lightningLord.state.hits = clamp(lightningLord.state.hits + extraHits, 3, 10);
          currentBattle.setActorSpeed(lightningLord, 60 + (lightningLord.state.hits - 3) * 10);
        }
      }
      return {};
    };

    battle.addActor(mainActor);
    addSupportActors(battle, context);
    return runBattle(battle);
  }

  function schemaFirefly() {
    return [
      {
        key: "completeCombustionStart",
        label: "开场处于完全燃烧",
        type: "checkbox",
        defaultValue: false,
        help: "按终结技已开启处理，开场获得 100% 行动提前、+60 速度，并生成 70 速倒计时。",
      },
    ];
  }

  function simulateFirefly(context) {
    const battle = createBattle(context);
    const settings = getSpecialSettings(context.specialSettings, "流萤");
    const actor = standardFocusActor(context, "流萤", context.focusRow);
    battle.addActor(actor);

    if (settings.completeCombustionStart) {
      battle.setActorSpeed(actor, actor.currentSpeed + 60);
      battle.immediateActor(actor, "完全燃烧");
      battle.addActor({
        id: "流萤-完全燃烧倒计时",
        label: "完全燃烧倒计时",
        roleName: "流萤",
        baseSpeed: 70,
        currentSpeed: 70,
        isFocusOwned: false,
        onTurn({ battle: currentBattle, actor: countdown }) {
          currentBattle.addTimeline(countdown.label, "完全燃烧结束");
          currentBattle.setActorSpeed(actor, Math.max(context.getEffectiveSpeed("流萤"), actor.currentSpeed - 60));
          countdown.active = false;
          return { resetAV: false };
        },
      });
    }

    addSupportActors(battle, context);
    return runBattle(battle);
  }

  function schemaXiadie() {
    return [
      {
        key: "dragonActive",
        label: "开场死龙在场",
        type: "checkbox",
        defaultValue: false,
        help: "按终结技/秘技后状态处理，死龙初始 165 速度，最多行动 3 次。",
      },
    ];
  }

  function simulateXiadie(context) {
    const battle = createBattle(context);
    const settings = getSpecialSettings(context.specialSettings, "遐蝶");
    const actor = standardFocusActor(context, "遐蝶", context.focusRow);
    const dragonRow =
      findRow(context.focusRole, (row) => row.sourceName === "燎尽黯泽的焰息") ||
      findRow(context.focusRole, (row) => row.rowType === "召唤物");
    battle.addActor(actor);

    if (settings.dragonActive && dragonRow) {
      const dragon = battle.addActor({
        id: "遐蝶-死龙",
        label: "死龙",
        roleName: "遐蝶",
        baseSpeed: 165,
        currentSpeed: 165,
        isFocusOwned: true,
        state: { remainingTurns: 3 },
        onTurn({ battle: currentBattle, actor: dragonActor }) {
          currentBattle.addDamage(dragonActor, dragonRow, 1, dragonRow.sourceName);
          currentBattle.addTimeline(dragonActor.label, `${dragonRow.sourceName}（剩余 ${dragonActor.state.remainingTurns - 1} 回合）`);
          dragonActor.state.remainingTurns -= 1;
          if (dragonActor.state.remainingTurns <= 0) {
            dragonActor.active = false;
          }
          return {};
        },
      });
      battle.immediateActor(dragon, "死龙行动提前100%");
    }

    addSupportActors(battle, context);
    return runBattle(battle);
  }

  function schemaBaiE() {
    return [
      {
        key: "khaslanaMode",
        label: "按卡厄斯兰那额外回合计算",
        type: "checkbox",
        defaultValue: false,
        help: "终结技变身后按 8 个额外回合模拟，速度按基础速度的 60% 处理。",
      },
    ];
  }

  function simulateBaiE(context) {
    const battle = createBattle(context);
    const settings = getSpecialSettings(context.specialSettings, "白厄");
    const finalHitRow = findRow(context.focusRole, (row) => row.rowType === "终结技");

    if (settings.khaslanaMode) {
      const transformed = battle.addActor({
        id: "白厄-卡厄斯兰那",
        label: "卡厄斯兰那",
        roleName: "白厄",
        baseSpeed: context.focusRole.baseSpeed * 0.6,
        currentSpeed: context.focusRole.baseSpeed * 0.6,
        isFocusOwned: true,
        state: { remainingTurns: 8 },
        onTurn({ battle: currentBattle, actor }) {
          currentBattle.addDamage(actor, context.focusRow, 1, context.focusRow.sourceName);
          actor.state.remainingTurns -= 1;
          currentBattle.addTimeline(actor.label, `额外回合，剩余 ${actor.state.remainingTurns}`);
          if (actor.state.remainingTurns === 0 && finalHitRow) {
            currentBattle.addDamage(actor, finalHitRow, 1, "最后一击");
            currentBattle.addTimeline(actor.label, "触发最后一击");
            actor.active = false;
          }
          return {};
        },
      });
      battle.immediateActor(transformed, "变身后开始额外回合");
      addSupportActors(battle, context);
      return runBattle(battle);
    }

    battle.addActor(standardFocusActor(context, "白厄", context.focusRow));
    addSupportActors(battle, context);
    return runBattle(battle);
  }

  function schemaYaoGuang() {
    return [
      {
        key: "ahaExtraTurns",
        label: "阿哈额外回合数",
        type: "number",
        min: 0,
        max: 3,
        step: 1,
        defaultValue: 0,
        help: "按终结技后额外回合次数近似；没有单独阿哈伤害条目时，仅统计行动次数。",
      },
    ];
  }

  function simulateYaoGuang(context) {
    const battle = createBattle(context);
    const settings = getSpecialSettings(context.specialSettings, "爻光");
    const actor = standardFocusActor(context, "爻光", context.focusRow);
    battle.addActor(actor);

    if (Number(settings.ahaExtraTurns || 0) > 0) {
      const aha = battle.addActor({
        id: "爻光-阿哈",
        label: "阿哈",
        roleName: "爻光",
        baseSpeed: actor.currentSpeed,
        currentSpeed: actor.currentSpeed,
        isFocusOwned: true,
        state: { remainingTurns: Number(settings.ahaExtraTurns || 0) },
        onTurn({ battle: currentBattle, actor: ahaActor }) {
          currentBattle.addTimeline(ahaActor.label, "额外回合（未单列伤害）");
          ahaActor.state.remainingTurns -= 1;
          if (ahaActor.state.remainingTurns <= 0) {
            ahaActor.active = false;
          }
          return {};
        },
      });
      battle.immediateActor(aha, "阿哈额外回合");
    }

    addSupportActors(battle, context);
    return runBattle(battle);
  }

  function simulateStandard(context, roleName) {
    const battle = createBattle(context);
    battle.addActor(standardFocusActor(context, roleName, context.focusRow));
    addSupportActors(battle, context);
    return runBattle(battle);
  }

  function addSupportActors(battle, context) {
    const focusActor = () => battle.actors.find((actor) => actor.id === `${context.focusRole.name}-main`) ||
      battle.actors.find((actor) => actor.roleName === context.focusRole.name && actor.isFocusOwned);

    context.teamRoles.forEach((roleName) => {
      if (roleName === context.focusRole.name) {
        return;
      }
      if (roleName === "花火") {
        battle.addActor(
          supportActor(context, roleName, ({ battle: currentBattle }) => {
            const target = focusActor();
            if (target) {
              currentBattle.advanceActor(target, 50);
              currentBattle.addTimeline("花火", `梦游鱼使 ${target.label} 行动提前 50%`);
            }
            return {};
          })
        );
      } else if (roleName === "布洛妮娅") {
        battle.addActor(
          supportActor(context, roleName, ({ battle: currentBattle }) => {
            const target = focusActor();
            if (target && target.label !== "布洛妮娅") {
              currentBattle.immediateActor(target, "作战再部署");
              currentBattle.addTimeline("布洛妮娅", `作战再部署使 ${target.label} 立即行动`);
            }
            return {};
          })
        );
      } else {
        battle.addActor(
          supportActor(context, roleName, ({ battle: currentBattle }) => {
            currentBattle.addTimeline(roleName, "常规行动");
            return {};
          })
        );
      }
    });
  }

  const schemas = {
    姬子: defaultSchema,
    希儿: schemaSeele,
    景元: schemaJingYuan,
    "丹恒•饮月": defaultSchema,
    卡芙卡: defaultSchema,
    流萤: schemaFirefly,
    遐蝶: schemaXiadie,
    白厄: schemaBaiE,
    花火: defaultSchema,
    "阮•梅": defaultSchema,
    布洛妮娅: defaultSchema,
    火花: defaultSchema,
    爻光: schemaYaoGuang,
  };

  const simulators = {
    姬子(context) {
      return simulateStandard(context, "姬子");
    },
    希儿(context) {
      return simulateSeele(context);
    },
    景元(context) {
      return simulateJingYuan(context);
    },
    "丹恒•饮月"(context) {
      return simulateStandard(context, "丹恒•饮月");
    },
    卡芙卡(context) {
      return simulateStandard(context, "卡芙卡");
    },
    流萤(context) {
      return simulateFirefly(context);
    },
    遐蝶(context) {
      return simulateXiadie(context);
    },
    白厄(context) {
      return simulateBaiE(context);
    },
    花火(context) {
      return simulateStandard(context, "花火");
    },
    "阮•梅"(context) {
      return simulateStandard(context, "阮•梅");
    },
    布洛妮娅(context) {
      return simulateStandard(context, "布洛妮娅");
    },
    火花(context) {
      return simulateStandard(context, "火花");
    },
    爻光(context) {
      return simulateYaoGuang(context);
    },
  };

  window.ROLE_ACTION_LOGIC = {
    getSchema(roleName) {
      return (schemas[roleName] || defaultSchema)();
    },
    simulate(context) {
      const simulator = simulators[context.focusRole.name] || ((ctx) => simulateStandard(ctx, ctx.focusRole.name));
      const result = simulator(context);
      return result;
    },
  };
})();
