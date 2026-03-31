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
    const isLastTripChecked = (vehicleId) => {
      if (typeof g.isDriverLastTripChecked === 'function') return !!g.isDriverLastTripChecked(Number(vehicleId || 0));
      return false;
    };
    const toNumber = (value) => {
      const n = Number(value);
      return Number.isFinite(n) ? n : 0;
    };
    const getRowPoint = (row) => {
      const lat = Number(row?.casts?.latitude ?? row?.latitude ?? row?.lat);
      const lng = Number(row?.casts?.longitude ?? row?.longitude ?? row?.lng);
      if (!Number.isFinite(lat) || !Number.isFinite(lng)) return null;
      return { lat, lng };
    };
    const getRowDistanceKm = (row) => {
      return Math.max(0, toNumber(row?.distance_km ?? row?.casts?.distance_km ?? 0));
    };
    const getRowTravelMinutes = (row) => {
      return Math.max(0, toNumber(row?.travel_minutes ?? row?.casts?.travel_minutes ?? 0));
    };
    const getRowDerivedSpeedKmh = (row) => {
      const km = Math.max(0, getRowDistanceKm(row));
      const minutes = Math.max(0, getRowTravelMinutes(row));
      if (km <= 0 || minutes <= 0) return 0;
      return (km / minutes) * 60;
    };
    const distanceBetweenPointsKm = (fromPoint, toPoint) => {
      if (!fromPoint || !toPoint) return 0;
      const rad = (deg) => (deg * Math.PI) / 180;
      const R = 6371;
      const dLat = rad(toPoint.lat - fromPoint.lat);
      const dLng = rad(toPoint.lng - fromPoint.lng);
      const lat1 = rad(fromPoint.lat);
      const lat2 = rad(toPoint.lat);
      const a = Math.sin(dLat / 2) ** 2 + Math.cos(lat1) * Math.cos(lat2) * Math.sin(dLng / 2) ** 2;
      const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
      return R * c;
    };
    const calcLegacyRoundTripSummary = (rows, vehicleId = 0) => {
      const ordered = Array.isArray(rows) ? rows.filter(Boolean) : [];
      if (!ordered.length) {
        return {
          outboundKm: 0,
          returnKm: 0,
          totalKm: 0,
          outboundMinutes: 0,
          returnMinutes: 0,
          totalMinutes: 0,
          legKm: [],
          legMinutes: [],
          model: 'legacy_segment_speed'
        };
      }

      let outboundKm = 0;
      let outboundMinutes = 0;
      const legKm = [];
      const legMinutes = [];
      let previousPoint = null;

      ordered.forEach((row, index) => {
        let km = 0;
        let minutes = 0;
        if (index === 0) {
          km = getRowDistanceKm(row);
          minutes = getRowTravelMinutes(row);
        } else {
          const currentPoint = getRowPoint(row);
          km = previousPoint && currentPoint
            ? distanceBetweenPointsKm(previousPoint, currentPoint)
            : getRowDistanceKm(row);
          const speed = getRowDerivedSpeedKmh(row);
          if (speed > 0) {
            minutes = (km / speed) * 60;
          } else {
            minutes = getRowTravelMinutes(row);
          }
        }
        km = Math.max(0, Number(km || 0));
        minutes = Math.max(0, Number(minutes || 0));
        legKm.push(Number(km.toFixed(1)));
        legMinutes.push(Math.round(minutes));
        outboundKm += km;
        outboundMinutes += minutes;
        previousPoint = getRowPoint(row) || previousPoint;
      });

      const lastRow = ordered[ordered.length - 1] || null;
      const returnKm = isLastTripChecked(vehicleId) ? 0 : getRowDistanceKm(lastRow);
      const returnMinutes = isLastTripChecked(vehicleId) ? 0 : getRowTravelMinutes(lastRow);
      const totalKm = outboundKm + returnKm;
      const totalMinutes = outboundMinutes + returnMinutes;

      return {
        outboundKm: Number(outboundKm.toFixed(1)),
        returnKm: Number(returnKm.toFixed(1)),
        totalKm: Number(totalKm.toFixed(1)),
        outboundMinutes: Math.round(outboundMinutes),
        returnMinutes: Math.round(returnMinutes),
        totalMinutes: Math.round(totalMinutes),
        legKm,
        legMinutes,
        model: 'legacy_segment_speed'
      };
    };
    const calcRouteKm = (rows, vehicleId = 0) => {
      return Number(calcLegacyRoundTripSummary(rows, vehicleId).totalKm || 0);
    };
    const calcRouteMinutes = (rows, vehicleId = 0) => {
      return Number(calcLegacyRoundTripSummary(rows, vehicleId).totalMinutes || 0);
    };

    const sortCandidateRows = (rows) => {
      const safeRows = Array.isArray(rows) ? rows.filter(Boolean) : [];
      if (typeof g.sortItemsByNearestRoute === 'function') {
        try {
          return g.sortItemsByNearestRoute(safeRows.map(entry => entry?.raw || entry));
        } catch (error) {}
      }
      return safeRows
        .map(entry => entry?.raw || entry)
        .sort((a, b) => Number(a?.distance_km || a?.casts?.distance_km || 0) - Number(b?.distance_km || b?.casts?.distance_km || 0) || Number(a?.id || 0) - Number(b?.id || 0));
    };

    const pre = Array.isArray(options.preAssignedAssignments) ? options.preAssignedAssignments : [];
    const preIds = new Set(pre.map(p => Number(p.item_id || 0)).filter(Boolean));

    const safeVehicles = (vehicles || []).filter(Boolean);
    const availableVehicles = safeVehicles.filter(v => (
      v.status === 'available' || (v.status === 'returning' && Number(v.return_in_minutes || 0) <= 10)
    ));
    const dispatchVehicles = availableVehicles.length ? availableVehicles : safeVehicles;
    if (!dispatchVehicles.length) {
      g.__THEMIS_LAST_OVERFLOW__ = {
        mode: 'fixed_append',
        vehicleCount: 0,
        actualGroupCount: 0,
        keptGroups: [],
        overflowGroups: [],
        evaluations: []
      };
      return [...pre];
    }

    const itemRecords = (items || []).filter(i => !preIds.has(Number(i?.id || 0))).map(item => {
      const area = normalizeArea(item.destination_area || item.cluster_area || item.casts?.area || '');
      return {
        raw: item,
        id: Number(item.id || 0),
        hour: Number(item.actual_hour ?? item.plan_hour ?? 0),
        travelMinutes: Number(item.travel_minutes || item.casts?.travel_minutes || 0),
        distanceKm: Number(item.distance_km || item.casts?.distance_km || 0),
        area,
        canonical: canonicalArea(area),
        group: displayGroup(area)
      };
    });

    const overallGroups = new Map();
    itemRecords.forEach(item => {
      if (!overallGroups.has(item.group)) overallGroups.set(item.group, []);
      overallGroups.get(item.group).push(item);
    });

    const orderedOverallGroups = [...overallGroups.entries()]
      .map(([group, rows]) => {
        const sortedDesc = rows.slice().sort((a, b) => Number(b.distanceKm || 0) - Number(a.distanceKm || 0) || Number(a.id || 0) - Number(b.id || 0));
        return {
          group,
          rows,
          anchor: sortedDesc[0] || null,
          count: rows.length
        };
      })
      .sort((a, b) => Number(b.anchor?.distanceKm || 0) - Number(a.anchor?.distanceKm || 0) || Number(b.count || 0) - Number(a.count || 0) || String(a.group || '').localeCompare(String(b.group || ''), 'ja'));

    const vehicleCount = Math.max(0, dispatchVehicles.length);
    const keptOverallGroups = orderedOverallGroups.slice(0, vehicleCount);
    const overflowOverallGroups = orderedOverallGroups.slice(vehicleCount);
    const keptGroupSet = new Set(keptOverallGroups.map(entry => entry.group));
    const overflowItemIdSet = new Set(
      overflowOverallGroups.flatMap(entry => entry.rows.map(row => Number(row.id || 0)).filter(Boolean))
    );

    const totalSeatCapacity = dispatchVehicles.reduce((sum, vehicle) => sum + Math.max(1, Number(vehicle?.seat_capacity || 4)), 0);
    const capacityOverflowCount = Math.max(0, itemRecords.length - totalSeatCapacity);
    const capacityOverflowItems = capacityOverflowCount > 0
      ? itemRecords
          .slice()
          .sort((a, b) => Number(a.distanceKm || 0) - Number(b.distanceKm || 0) || Number(a.id || 0) - Number(b.id || 0))
          .slice(0, capacityOverflowCount)
      : [];
    const capacityOverflowIdSet = new Set(capacityOverflowItems.map(item => Number(item.id || 0)).filter(Boolean));
    const dispatchableItemRecords = itemRecords.filter(item => !capacityOverflowIdSet.has(Number(item.id || 0)));

    const evaluations = [];

    const stateMap = new Map();
    const ensureState = (vehicleId, hour) => {
      const key = `${Number(vehicleId || 0)}__${Number(hour || 0)}`;
      if (!stateMap.has(key)) {
        stateMap.set(key, {
          fixedRows: [],
          addedRows: [],
          fixedGroups: new Set(),
          totalCount: 0
        });
      }
      return stateMap.get(key);
    };

    const pushAssignment = (results, item, vehicleId) => {
      results.push({
        item_id: Number(item.id || 0),
        vehicle_id: Number(vehicleId || 0),
        actual_hour: Number(item.hour || 0),
        distance_km: Number(item.distanceKm || 0)
      });
    };

    pre.forEach(p => {
      const rawItem = (items || []).find(x => Number(x?.id || 0) === Number(p?.item_id || 0));
      if (!rawItem) return;
      const item = {
        raw: rawItem,
        id: Number(rawItem.id || 0),
        hour: Number(rawItem.actual_hour ?? rawItem.plan_hour ?? p.actual_hour ?? 0),
        travelMinutes: Number(rawItem.travel_minutes || rawItem.casts?.travel_minutes || p.travel_minutes || 0),
        distanceKm: Number(rawItem.distance_km || rawItem.casts?.distance_km || p.distance_km || 0),
        area: normalizeArea(rawItem.destination_area || rawItem.cluster_area || rawItem.casts?.area || p.destination_area || ''),
        canonical: canonicalArea(rawItem.destination_area || rawItem.cluster_area || rawItem.casts?.area || p.destination_area || ''),
        group: displayGroup(rawItem.destination_area || rawItem.cluster_area || rawItem.casts?.area || p.destination_area || '')
      };
      const state = ensureState(Number(p.vehicle_id || 0), item.hour);
      state.fixedRows.push(item);
      if (item.group) state.fixedGroups.add(item.group);
      state.totalCount += 1;
    });

    const groupedByHour = new Map();
    dispatchableItemRecords.forEach(item => {
      if (!groupedByHour.has(item.hour)) groupedByHour.set(item.hour, []);
      groupedByHour.get(item.hour).push(item);
    });

    const vehicleBaseScore = (vehicle, item) => {
      const vehicleArea = normalizeArea(vehicle?.vehicle_area || '');
      const homeArea = normalizeArea(vehicle?.home_area || '');
      const itemArea = item.area;

      let score = 0;
      score += areaAffinity(vehicleArea, itemArea) * 2.0;
      score += areaAffinity(homeArea, itemArea) * 2.2;
      score += Math.max(0, directionAffinity(vehicleArea, itemArea)) * 1.5;
      score += Math.max(0, directionAffinity(homeArea, itemArea)) * 1.7;
      score += strictHome(itemArea, homeArea) * 2.0;
      if (isLastTripChecked(vehicle?.id) && displayGroup(homeArea) === item.group) score += 120;
      return score;
    };

    const chooseDedicatedVehicles = (keptHourItems) => {
      const groups = new Map();
      keptHourItems.forEach(item => {
        if (!groups.has(item.group)) groups.set(item.group, []);
        groups.get(item.group).push(item);
      });

      const orderedGroups = [...groups.entries()]
        .map(([group, rows]) => {
          const sortedDesc = rows.slice().sort((a, b) => Number(b.distanceKm || 0) - Number(a.distanceKm || 0) || Number(a.id || 0) - Number(b.id || 0));
          return { group, rows, anchor: sortedDesc[0] || null, count: rows.length };
        })
        .sort((a, b) => Number(b.anchor?.distanceKm || 0) - Number(a.anchor?.distanceKm || 0) || Number(b.count || 0) - Number(a.count || 0));

      const dedicated = new Map();
      const used = new Set();

      orderedGroups.forEach(entry => {
        let best = null;
        dispatchVehicles.forEach(vehicle => {
          const vid = Number(vehicle?.id || 0);
          if (!vid || used.has(vid)) return;
          let score = vehicleBaseScore(vehicle, entry.anchor || entry.rows[0] || {});
          score -= Number(entry.count || 0) * 2;
          if (!best || score > best.score) best = { vehicleId: vid, score };
        });
        if (best) {
          used.add(best.vehicleId);
          dedicated.set(entry.group, best.vehicleId);
        }
      });

      return dedicated;
    };

    const results = [...pre];
    const hourKeys = [...groupedByHour.keys()].sort((a, b) => a - b);

    hourKeys.forEach(hour => {
      const hourItems = groupedByHour.get(hour) || [];
      const fixedHourItems = hourItems.filter(item => keptGroupSet.has(item.group));
      const overflowHourItems = hourItems.filter(item => overflowItemIdSet.has(item.id));

      const dedicatedVehicles = chooseDedicatedVehicles(fixedHourItems);

      const fixedGroups = new Map();
      fixedHourItems.forEach(item => {
        if (!fixedGroups.has(item.group)) fixedGroups.set(item.group, []);
        fixedGroups.get(item.group).push(item);
      });

      [...fixedGroups.entries()].forEach(([group, rows]) => {
        const vehicleId = Number(dedicatedVehicles.get(group) || 0);
        if (!vehicleId) return;
        const state = ensureState(vehicleId, hour);
        rows.slice()
          .sort((a, b) => Number(a.distanceKm || 0) - Number(b.distanceKm || 0) || Number(a.id || 0) - Number(b.id || 0))
          .forEach(item => {
            pushAssignment(results, item, vehicleId);
            state.fixedRows.push(item);
            state.fixedGroups.add(item.group);
            state.totalCount += 1;
          });
      });

      overflowHourItems
        .slice()
        .sort((a, b) => Number(a.distanceKm || 0) - Number(b.distanceKm || 0) || Number(a.id || 0) - Number(b.id || 0))
        .forEach(item => {
          let best = null;
          const itemEval = {
            itemId: Number(item.id || 0),
            castName: String(item.raw?.casts?.name || item.raw?.name || ''),
            group: item.group,
            area: item.area,
            distanceKm: Number(item.distanceKm || 0),
            hour: Number(item.hour || 0),
            candidates: []
          };

          dispatchVehicles.forEach(vehicle => {
            const vid = Number(vehicle?.id || 0);
            if (!vid) return;
            const seatCapacity = Math.max(1, Number(vehicle?.seat_capacity || 4));
            const state = ensureState(vid, hour);
            if (state.totalCount >= seatCapacity) {
              itemEval.candidates.push({
                vehicleId: vid,
                vehicleName: String(vehicle?.driver_name || vehicle?.plate_number || `車両${vid}`),
                excluded: true,
                reason: 'seat_full'
              });
              return;
            }

            let maxReturnMinutes = 0;
            let totalRoundTripKm = 0;
            let targetRouteMinutes = 0;
            let targetRouteKm = 0;
            let targetOutboundMinutes = 0;
            let targetReturnMinutes = 0;
            let targetOutboundKm = 0;
            let targetReturnKm = 0;

            dispatchVehicles.forEach(worldVehicle => {
              const worldVid = Number(worldVehicle?.id || 0);
              if (!worldVid) return;
              const worldState = ensureState(worldVid, hour);
              const worldRows = [...worldState.fixedRows, ...worldState.addedRows];
              if (worldVid === vid) worldRows.push(item);
              const combinedRows = sortCandidateRows(worldRows);
              const worldSummary = calcLegacyRoundTripSummary(combinedRows, worldVid);
              const worldMinutes = Number(worldSummary.totalMinutes || 0);
              const worldKm = Number(worldSummary.totalKm || 0);

              if (worldMinutes > maxReturnMinutes) maxReturnMinutes = worldMinutes;
              totalRoundTripKm += worldKm;

              if (worldVid === vid) {
                targetRouteMinutes = worldMinutes;
                targetRouteKm = worldKm;
                targetOutboundMinutes = Number(worldSummary.outboundMinutes || 0);
                targetReturnMinutes = Number(worldSummary.returnMinutes || 0);
                targetOutboundKm = Number(worldSummary.outboundKm || 0);
                targetReturnKm = Number(worldSummary.returnKm || 0);
              }
            });

            totalRoundTripKm = Number(totalRoundTripKm.toFixed(1));
            const fixedGroupsList = [...state.fixedGroups.values()];
            const candidate = {
              vehicleId: vid,
              vehicleName: String(vehicle?.driver_name || vehicle?.plate_number || `車両${vid}`),
              maxReturnMinutes,
              totalRoundTripKm,
              routeMinutes: targetRouteMinutes,
              routeKm: targetRouteKm,
              outboundMinutes: targetOutboundMinutes,
              returnMinutes: targetReturnMinutes,
              outboundKm: targetOutboundKm,
              returnKm: targetReturnKm,
              seatCountAfter: state.totalCount + 1,
              fixedGroups: fixedGroupsList,
              model: 'legacy_segment_speed'
            };
            itemEval.candidates.push(candidate);

            const better =
              !best ||
              maxReturnMinutes < best.maxReturnMinutes ||
              (maxReturnMinutes === best.maxReturnMinutes && totalRoundTripKm < best.totalRoundTripKm) ||
              (maxReturnMinutes === best.maxReturnMinutes && totalRoundTripKm === best.totalRoundTripKm && (state.totalCount + 1) < best.seatCountAfter) ||
              (maxReturnMinutes === best.maxReturnMinutes && totalRoundTripKm === best.totalRoundTripKm && (state.totalCount + 1) === best.seatCountAfter && vid < best.vehicleId);

            if (better) {
              best = {
                vehicleId: vid,
                maxReturnMinutes,
                totalRoundTripKm,
                seatCountAfter: state.totalCount + 1,
                routeMinutes: targetRouteMinutes,
                routeKm: targetRouteKm
              };
            }
          });

          if (best) {
            const state = ensureState(best.vehicleId, hour);
            state.addedRows.push(item);
            state.totalCount += 1;
            pushAssignment(results, item, best.vehicleId);
            itemEval.selectedVehicleId = best.vehicleId;
            itemEval.selectedMaxReturnMinutes = best.maxReturnMinutes;
            itemEval.selectedTotalRoundTripKm = best.totalRoundTripKm;
            itemEval.selectedSeatCountAfter = best.seatCountAfter;
            itemEval.selectedRouteMinutes = best.routeMinutes;
            itemEval.selectedRouteKm = best.routeKm;
          } else {
            itemEval.unassigned = true;
          }

          evaluations.push(itemEval);
        });
    });

    const assignedIds = new Set(results.map(row => Number(row.item_id || 0)).filter(Boolean));
    const unresolvedOverflowGroups = overflowOverallGroups
      .map(entry => {
        const unresolvedIds = entry.rows.map(row => Number(row.id || 0)).filter(id => !assignedIds.has(id));
        return unresolvedIds.length ? {
          group: entry.group,
          distanceKm: Number(entry.anchor?.distanceKm || 0),
          count: unresolvedIds.length,
          itemIds: unresolvedIds
        } : null;
      })
      .filter(Boolean);

    g.__THEMIS_LAST_OVERFLOW__ = {
      mode: 'fixed_append',
      vehicleCount,
      actualGroupCount: orderedOverallGroups.length,
      keptGroups: keptOverallGroups.map(entry => ({
        group: entry.group,
        distanceKm: Number(entry.anchor?.distanceKm || 0),
        count: Number(entry.count || 0)
      })),
      overflowGroups: unresolvedOverflowGroups,
      evaluations,
      totalSeatCapacity,
      totalCastCount: itemRecords.length,
      capacityOverflowCount,
      capacityOverflowItems: capacityOverflowItems.map(item => ({
        itemId: Number(item.id || 0),
        hour: Number(item.hour || 0),
        castName: String(item.raw?.casts?.name || item.raw?.name || ''),
        distanceKm: Number(item.distanceKm || 0),
        group: item.group,
        area: item.area,
        reason: 'capacity_overflow'
      }))
    };

    return results;
  }
};
