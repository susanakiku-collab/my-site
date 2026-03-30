window.DispatchCore = {
  optimizeAssignments: function (items, vehicles, monthlyMap, options = {}) {
    const g = globalThis;

    const normalizeArea = (area) => {
      if (typeof g.normalizeAreaLabel === 'function') return g.normalizeAreaLabel(area || '');
      return String(area || '').trim();
    };
    const canonicalArea = (area) => {
      const a = normalizeArea(area);
      if (typeof g.getCanonicalArea === 'function') return g.getCanonicalArea(a) || a;
      return a;
    };
    const displayGroup = (area) => {
      const a = normalizeArea(area);
      if (typeof g.getAreaDisplayGroup === 'function') return g.getAreaDisplayGroup(a) || a;
      return a;
    };
    const areaAffinity = (a, b) => {
      if (typeof g.getAreaAffinityScore === 'function') return Number(g.getAreaAffinityScore(a, b) || 0);
      return canonicalArea(a) === canonicalArea(b) ? 100 : 0;
    };
    const directionAffinity = (a, b) => {
      if (typeof g.getDirectionAffinityScore === 'function') return Number(g.getDirectionAffinityScore(a, b) || 0);
      return canonicalArea(a) === canonicalArea(b) ? 100 : 0;
    };
    const strictHome = (area, home) => {
      if (typeof g.getStrictHomeCompatibilityScore === 'function') return Number(g.getStrictHomeCompatibilityScore(area, home) || 0);
      return canonicalArea(area) === canonicalArea(home) ? 100 : 0;
    };
    const isHardReverseRoute = (a, b) => {
      if (typeof g.isHardReverseMixForRoute === 'function') return !!g.isHardReverseMixForRoute(a, b);
      return false;
    };
    const isHardReverseHome = (area, home) => {
      if (typeof g.isHardReverseForHome === 'function') return !!g.isHardReverseForHome(area, home);
      return false;
    };
    const isLastTripChecked = (vehicleId) => {
      if (typeof g.isDriverLastTripChecked === 'function') return !!g.isDriverLastTripChecked(Number(vehicleId || 0));
      return false;
    };

    const pre = Array.isArray(options.preAssignedAssignments) ? options.preAssignedAssignments : [];
    const preIds = new Set(pre.map(p => Number(p.item_id)));

    const safeVehicles = (vehicles || []).filter(Boolean);
    const availableVehicles = safeVehicles.filter(v => (
      v.status === 'available' || (v.status === 'returning' && Number(v.return_in_minutes || 0) <= 10)
    ));

    const vehicleLoad = {};
    const vehicleHourAssignments = new Map();
    const vehicleHourGroups = new Map();

    const ensureHourState = (vehicleId, hour) => {
      const key = `${Number(vehicleId || 0)}__${Number(hour || 0)}`;
      if (!vehicleHourAssignments.has(key)) vehicleHourAssignments.set(key, []);
      if (!vehicleHourGroups.has(key)) vehicleHourGroups.set(key, new Set());
      return key;
    };

    availableVehicles.forEach(v => {
      vehicleLoad[Number(v.id)] = { minutes: 0, count: 0 };
    });

    pre.forEach(p => {
      const vid = Number(p.vehicle_id);
      if (!vehicleLoad[vid]) vehicleLoad[vid] = { minutes: 0, count: 0 };
      vehicleLoad[vid].minutes += Number(p.travel_minutes || p.distance_km || 0);
      vehicleLoad[vid].count += 1;
      const hourKey = ensureHourState(vid, p.actual_hour);
      const area = normalizeArea(p.destination_area || p.area || p.cluster_area || p.casts?.area || '');
      if (area) {
        vehicleHourAssignments.get(hourKey).push(area);
        vehicleHourGroups.get(hourKey).add(displayGroup(area));
      }
    });

    const newItems = (items || []).filter(i => !preIds.has(Number(i.id))).map(item => {
      const area = normalizeArea(item.destination_area || item.cluster_area || item.casts?.area || '');
      return {
        raw: item,
        id: Number(item.id),
        hour: Number(item.actual_hour ?? item.plan_hour ?? 0),
        travelMinutes: Number(item.travel_minutes || 0),
        distanceKm: Number(item.distance_km || item.casts?.distance_km || 0),
        area,
        canonical: canonicalArea(area),
        group: displayGroup(area)
      };
    });

    const groupedByHour = new Map();
    newItems.forEach(item => {
      if (!groupedByHour.has(item.hour)) groupedByHour.set(item.hour, []);
      groupedByHour.get(item.hour).push(item);
    });

    const results = [];

    const vehicleAreaScore = (vehicle, item) => {
      const vehicleArea = normalizeArea(vehicle?.vehicle_area || '');
      const homeArea = normalizeArea(vehicle?.home_area || '');
      const itemArea = item.area;
      const itemCanonical = item.canonical;
      const itemGroup = item.group;

      let score = 0;

      const vehicleCanonical = canonicalArea(vehicleArea);
      const homeCanonical = canonicalArea(homeArea);
      const vehicleGroup = displayGroup(vehicleArea);
      const homeGroup = displayGroup(homeArea);

      if (vehicleCanonical && itemCanonical && vehicleCanonical === itemCanonical) score += 320;
      if (homeCanonical && itemCanonical && homeCanonical === itemCanonical) score += 340;
      if (vehicleGroup && itemGroup && vehicleGroup === itemGroup) score += 220;
      if (homeGroup && itemGroup && homeGroup === itemGroup) score += 240;

      score += areaAffinity(vehicleArea, itemArea) * 2.0;
      score += areaAffinity(homeArea, itemArea) * 2.2;
      score += Math.max(0, directionAffinity(vehicleArea, itemArea)) * 1.7;
      score += Math.max(0, directionAffinity(homeArea, itemArea)) * 1.9;
      score += strictHome(itemArea, homeArea) * 2.8;

      if (isHardReverseHome(itemArea, homeArea)) score -= 4200;
      else {
        const homeDir = directionAffinity(itemArea, homeArea);
        if (homeDir < 0) score += homeDir * 26;
      }

      return score;
    };

    const chooseDedicatedVehicles = (hourItems) => {
      const groups = new Map();
      hourItems.forEach(item => {
        if (!groups.has(item.group)) groups.set(item.group, []);
        groups.get(item.group).push(item);
      });
      // 方面数が可能台数を超えても、遠い方面から可能台数ぶんは必ず先に固定する。
      // 4台5方面なら、最遠4方面を dedicated 対象にし、残り1方面だけを後段の例外処理へ回す。
      if (groups.size === 0) return new Map();

      const orderedGroups = [...groups.entries()]
        .map(([group, rows]) => ({
          group,
          rows: rows.slice().sort((a, b) => Number(b.distanceKm || 0) - Number(a.distanceKm || 0)),
          size: rows.length,
          anchor: rows.slice().sort((a, b) => Number(b.distanceKm || 0) - Number(a.distanceKm || 0))[0]
        }))
        .sort((a, b) => Number(b.anchor?.distanceKm || 0) - Number(a.anchor?.distanceKm || 0) || b.size - a.size);

      const dedicated = new Map();
      const used = new Set();
      orderedGroups.forEach(entry => {
        let best = null;
        availableVehicles.forEach(vehicle => {
          const vid = Number(vehicle.id || 0);
          if (used.has(vid)) return;
          let score = vehicleAreaScore(vehicle, entry.anchor);
          if (isLastTripChecked(vid) && displayGroup(vehicle?.home_area || '') === entry.group) score += 260;
          if (score > (best?.score ?? -Infinity)) best = { vehicle, score };
        });
        if (best?.vehicle) {
          const vid = Number(best.vehicle.id || 0);
          used.add(vid);
          dedicated.set(entry.group, vid);
        }
      });
      return dedicated;
    };

    const hourKeys = [...groupedByHour.keys()].sort((a, b) => a - b);

    hourKeys.forEach(hour => {
      const hourItems = groupedByHour.get(hour) || [];
      const dedicatedVehicles = chooseDedicatedVehicles(hourItems);

      const orderedItems = hourItems.slice().sort((a, b) => {
        const distDiff = Number(b.distanceKm || 0) - Number(a.distanceKm || 0);
        if (distDiff !== 0) return distDiff;
        return Number(b.travelMinutes || 0) - Number(a.travelMinutes || 0);
      });

      orderedItems.forEach(item => {
        if (options.mode !== 'last_trip' && Number(item.travelMinutes || 0) > 60) {
          return;
        }

        let best = null;

        availableVehicles.forEach(vehicle => {
          const vid = Number(vehicle.id || 0);
          const load = vehicleLoad[vid] || { minutes: 0, count: 0 };
          const hourKey = ensureHourState(vid, hour);
          const sameHourAreas = vehicleHourAssignments.get(hourKey) || [];
          const sameHourGroups = vehicleHourGroups.get(hourKey) || new Set();
          const seatCapacity = Math.max(1, Number(vehicle?.seat_capacity || 4));
          if (sameHourAreas.length >= seatCapacity) return;

          let score = vehicleAreaScore(vehicle, item);

          // 方面グループ優先
          if (dedicatedVehicles.has(item.group)) {
            const dedicatedVid = Number(dedicatedVehicles.get(item.group) || 0);
            if (vid === dedicatedVid) score += 520;
            else if (dedicatedVehicles.size > 0 && dedicatedVid !== vid) score -= 260;
          }

          // 同方向まとめ / 逆方向排除
          sameHourAreas.forEach(existingArea => {
            const existingCanonical = canonicalArea(existingArea);
            const existingGroup = displayGroup(existingArea);
            if (existingCanonical && item.canonical && existingCanonical === item.canonical) score += 360;
            if (existingGroup && item.group && existingGroup === item.group) score += 260;

            const affinity = areaAffinity(existingArea, item.area);
            const direction = directionAffinity(existingArea, item.area);
            score += affinity * 2.0;
            score += Math.max(0, direction) * 1.8;

            if (isHardReverseRoute(existingArea, item.area) || direction <= -38) {
              score -= 4800;
            } else if (direction < 0) {
              score += direction * 34;
            }
          });

          if (sameHourGroups.size > 0 && !sameHourGroups.has(item.group)) {
            score -= 180;
          }

          // ラスト便補正
          const checked = isLastTripChecked(vid);
          if (checked) {
            const homeGroup = displayGroup(vehicle?.home_area || '');
            const homeCanonical = canonicalArea(vehicle?.home_area || '');
            if (homeCanonical && homeCanonical === item.canonical) score += 520;
            if (homeGroup && homeGroup === item.group) score += 300;
            if (isHardReverseHome(item.area, vehicle?.home_area || '')) score -= 5200;
          } else {
            // ラスト便対象ドライバーがいれば、そのホーム方面案件はそちらへ寄せやすくする
            const anyChecked = availableVehicles.some(v => isLastTripChecked(v.id));
            if (anyChecked) {
              const checkedHomeMatch = availableVehicles.some(v => isLastTripChecked(v.id) && displayGroup(v?.home_area || '') === item.group);
              if (checkedHomeMatch) score -= 120;
            }
          }

          // 負荷バランス
          const month = monthlyMap?.get?.(vid) || { totalDistance: 0, avgDistance: 0, workedDays: 0 };
          score -= Number(load.minutes || 0) * 0.8;
          score -= Number(load.count || 0) * 28;
          score -= Number(month.totalDistance || 0) * 0.12;
          score -= Number(month.avgDistance || 0) * 9.5;

          if (score > (best?.score ?? -Infinity)) {
            best = { vehicle, score };
          }
        });

        if (best?.vehicle) {
          const vehicle = best.vehicle;
          const vid = Number(vehicle.id || 0);
          const hourKey = ensureHourState(vid, hour);
          results.push({
            item_id: item.id,
            vehicle_id: vid,
            actual_hour: item.hour,
            distance_km: item.distanceKm
          });
          vehicleLoad[vid] = vehicleLoad[vid] || { minutes: 0, count: 0 };
          vehicleLoad[vid].minutes += Number(item.travelMinutes || 0);
          vehicleLoad[vid].count += 1;
          vehicleHourAssignments.get(hourKey).push(item.area);
          vehicleHourGroups.get(hourKey).add(item.group);
        }
      });
    });

    return [...pre, ...results];
  }
};
