window.DispatchCore = {
  optimizeAssignments: function (items, vehicles, monthlyMap, options = {}) {

    const pre = options.preAssignedAssignments || [];
    const preIds = new Set(pre.map(p => Number(p.item_id)));

    const vehicleLoad = {};
    (vehicles || []).forEach(v => {
      vehicleLoad[Number(v.id)] = { minutes: 0, count: 0 };
    });

    // ===== 既存負荷 =====
    pre.forEach(p => {
      const vid = Number(p.vehicle_id);
      if (!vehicleLoad[vid]) vehicleLoad[vid] = { minutes: 0, count: 0 };
      vehicleLoad[vid].minutes += Number(p.distance_km || 0);
      vehicleLoad[vid].count += 1;
    });

    const newItems = (items || []).filter(i => !preIds.has(Number(i.id)));

    const results = [];

    newItems.forEach(item => {

      // ===== 60分ルール =====
      if (options.mode !== "last_trip" && Number(item.travel_minutes) > 60) {
        return; // 通常便は除外
      }

      let bestVehicle = null;
      let bestTime = Infinity;

      (vehicles || []).forEach(v => {
        const vid = Number(v.id);

        // ===== 横ズレ制約（テスト用フック）=====
        if (typeof v.test_side_km === "number" && v.test_side_km > 3) {
          return;
        }

        // ===== 使用可能車両 =====
        if (
          v.status !== "available" &&
          !(v.status === "returning" && Number(v.return_in_minutes || 0) <= 10)
        ) return;

        const load = vehicleLoad[vid] || { minutes: 0, count: 0 };

        // ===== コア：時間優先 =====
        let estimatedTime = Number(load.minutes || 0) + Number(item.travel_minutes || 0);

        // ===== 方向補正（弱）=====
        if (v.vehicle_area === item.destination_area) {
          estimatedTime -= 5;
        }

        // ===== ラスト便：帰宅方向優先 =====
        if (options.mode === "last_trip") {
          if (v.home_area === item.destination_area) {
            estimatedTime -= 10;
          }
        }

        if (estimatedTime < bestTime) {
          bestTime = estimatedTime;
          bestVehicle = v;
        }
      });

      if (bestVehicle) {
        const vid = Number(bestVehicle.id);

        results.push({
          item_id: Number(item.id),
          vehicle_id: vid,
          actual_hour: item.actual_hour,
          distance_km: item.distance_km
        });

        if (!vehicleLoad[vid]) vehicleLoad[vid] = { minutes: 0, count: 0 };
        vehicleLoad[vid].minutes += Number(item.travel_minutes || 0);
        vehicleLoad[vid].count += 1;
      }

    });

    return [...pre, ...results];
  }
};