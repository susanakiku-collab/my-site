(function (global) {
  'use strict';

  const VERSION = 'dispatchcore-2026-04-07-normal-merge-waiting';
  const AREA_MEMBER_ANGLE_THRESHOLD = 30;
  const SAME_DIRECTION_MERGE_ANGLE_THRESHOLD = 18;
  const SAME_DIRECTION_MERGE_POINT_DISTANCE_KM = 18;
  const DISPATCH_DEBUG = Boolean(global.DISPATCH_DEBUG);

  function debugLog(...args) {
    if (!DISPATCH_DEBUG) return;
    console.log(...args);
  }

  function debugWarn(...args) {
    if (!DISPATCH_DEBUG) return;
    console.warn(...args);
  }

  function toNumber(value, fallback = 0) {
    const num = Number(value);
    return Number.isFinite(num) ? num : fallback;
  }

  function normalizeAngle(angle) {
    const normalized = angle % 360;
    return normalized < 0 ? normalized + 360 : normalized;
  }

  function angleDiff(a, b) {
    const diff = Math.abs(normalizeAngle(a) - normalizeAngle(b));
    return diff > 180 ? 360 - diff : diff;
  }

  function haversineKm(a, b) {
    const lat1 = toNumber(a?.lat, NaN);
    const lng1 = toNumber(a?.lng, NaN);
    const lat2 = toNumber(b?.lat, NaN);
    const lng2 = toNumber(b?.lng, NaN);
    if (![lat1, lng1, lat2, lng2].every(Number.isFinite)) return Infinity;

    const R = 6371;
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLng = (lng2 - lng1) * Math.PI / 180;
    const rad1 = lat1 * Math.PI / 180;
    const rad2 = lat2 * Math.PI / 180;
    const sinLat = Math.sin(dLat / 2);
    const sinLng = Math.sin(dLng / 2);
    const h = sinLat * sinLat + Math.cos(rad1) * Math.cos(rad2) * sinLng * sinLng;
    return 2 * R * Math.asin(Math.sqrt(h));
  }

  function calcAngleFromOrigin(origin, point) {
    const dx = point.lng - origin.lng;
    const dy = point.lat - origin.lat;
    return normalizeAngle(Math.atan2(dy, dx) * 180 / Math.PI);
  }

  function getOrigin() {
    const cfg = global.APP_CONFIG || {};
    const lat = toNumber(cfg.ORIGIN_LAT, NaN);
    const lng = toNumber(cfg.ORIGIN_LNG, NaN);
    return { lat, lng, name: cfg.ORIGIN_LABEL || '起点' };
  }

  function getItemPoint(item) {
    const lat = toNumber(item?.casts?.latitude ?? item?.latitude ?? item?.lat, NaN);
    const lng = toNumber(item?.casts?.longitude ?? item?.longitude ?? item?.lng, NaN);
    if (!Number.isFinite(lat) || !Number.isFinite(lng)) return null;
    return { lat, lng };
  }

  function getItemDistanceKm(item, origin, point) {
    const stored = toNumber(item?.distance_km ?? item?.casts?.distance_km, NaN);
    if (Number.isFinite(stored) && stored > 0) return stored;
    if (!point) return Infinity;
    return haversineKm(origin, point);
  }

  function normalizeItem(item, origin) {
    const point = getItemPoint(item);
    const distanceKm = getItemDistanceKm(item, origin, point);
    const angleFromOrigin = point ? calcAngleFromOrigin(origin, point) : null;
    return {
      itemId: Number(item?.id || 0),
      castId: Number(item?.cast_id || 0),
      actualHour: Number(item?.actual_hour ?? item?.plan_hour ?? 0),
      name: String(item?.casts?.name || item?.name || `item_${item?.id || ''}` || '').trim() || '-',
      area: String(item?.destination_area || item?.planned_area || item?.casts?.area || '').trim(),
      distanceKm,
      angleFromOrigin,
      point,
      raw: item
    };
  }

  function normalizeVehicle(vehicle, monthlyMap) {
    const vehicleId = Number(vehicle?.id || 0);
    const monthly = monthlyMap?.get(vehicleId) || {};
    const avgDistance = toNumber(monthly?.avgDistance ?? monthly?.avg_distance ?? monthly?.averageDistance, 0);
    const totalDistance = toNumber(monthly?.totalDistance ?? monthly?.total_distance, 0);
    const workedDays = Math.max(0, Math.trunc(toNumber(monthly?.workedDays ?? monthly?.worked_days, 0)));
    const homeLat = toNumber(vehicle?.home_lat, NaN);
    const homeLng = toNumber(vehicle?.home_lng, NaN);
    const isLastTrip = typeof global.isDriverLastTripChecked === 'function'
      ? Boolean(global.isDriverLastTripChecked(vehicleId))
      : Boolean(vehicle?.is_last_trip || vehicle?.isLastTrip);
    return {
      vehicleId,
      driverName: String(vehicle?.driver_name || vehicle?.name || vehicle?.plate_number || '').trim() || '-',
      capacity: Math.max(0, Math.trunc(toNumber(vehicle?.seat_capacity ?? vehicle?.capacity, 0))),
      avgDistance,
      totalDistance,
      workedDays,
      todayDistance: 0,
      score: avgDistance,
      homeArea: String(vehicle?.home_area || '').trim(),
      homePoint: Number.isFinite(homeLat) && Number.isFinite(homeLng) ? { lat: homeLat, lng: homeLng } : null,
      isLastTrip,
      raw: vehicle
    };
  }

  function byDistanceDesc(a, b) {
    if (b.distanceKm !== a.distanceKm) return b.distanceKm - a.distanceKm;
    return String(a.name || '').localeCompare(String(b.name || ''), 'ja');
  }

  function byDistanceAsc(a, b) {
    if (a.distanceKm !== b.distanceKm) return a.distanceKm - b.distanceKm;
    return String(a.name || '').localeCompare(String(b.name || ''), 'ja');
  }

  function buildAxis(anchor, index, vehicle) {
    return {
      axisId: `axis_${index + 1}`,
      vehicleId: vehicle.vehicleId,
      driverName: vehicle.driverName,
      capacity: vehicle.capacity,
      anchorItemId: anchor.itemId,
      anchorName: anchor.name,
      anchorAngle: anchor.angleFromOrigin,
      anchorDistanceKm: anchor.distanceKm,
      members: [anchor],
      pending: [],
      overflowed: []
    };
  }

  function getPrimaryCandidateAxes(item, axes, threshold) {
    return axes
      .map(axis => ({ axis, diff: angleDiff(item.angleFromOrigin, axis.anchorAngle) }))
      .filter(entry => Number.isFinite(entry.diff) && entry.diff <= threshold)
      .sort((a, b) => {
        if (a.diff !== b.diff) return a.diff - b.diff;
        if (b.axis.anchorDistanceKm !== a.axis.anchorDistanceKm) return b.axis.anchorDistanceKm - a.axis.anchorDistanceKm;
        return String(a.axis.axisId).localeCompare(String(b.axis.axisId), 'ja');
      });
  }

  function assignStagewiseCoveredIds(people, axes, threshold) {
    const coveredIds = new Set(axes.map(axis => axis.anchorItemId));
    people.forEach(item => {
      if (coveredIds.has(item.itemId)) return;
      const matches = getPrimaryCandidateAxes(item, axes, threshold);
      if (!matches.length) return;
      coveredIds.add(item.itemId);
    });
    return coveredIds;
  }

  function selectFallbackAxisCandidate(sorted, selectedAnchorIds, axes) {
    const remaining = sorted.filter(item => !selectedAnchorIds.has(item.itemId));
    if (!remaining.length) return null;
    if (!axes.length) return remaining[0] || null;

    return remaining
      .map(item => {
        const minAngleDiff = axes.reduce((min, axis) => {
          const diff = angleDiff(item.angleFromOrigin, axis.anchorAngle);
          return Math.min(min, diff);
        }, Infinity);
        return { item, minAngleDiff };
      })
      .sort((a, b) => {
        if (b.minAngleDiff !== a.minAngleDiff) return b.minAngleDiff - a.minAngleDiff;
        if (b.item.distanceKm !== a.item.distanceKm) return b.item.distanceKm - a.item.distanceKm;
        return String(a.item.name || '').localeCompare(String(b.item.name || ''), 'ja');
      })[0]?.item || null;
  }

  function selectFixedAxes(people, vehicles, threshold) {
    const sorted = [...people].sort(byDistanceDesc);
    const axes = [];
    if (!sorted.length || !vehicles.length) return axes;

    const selectedAnchorIds = new Set();
    let coveredIds = new Set();

    for (let i = 0; i < vehicles.length; i += 1) {
      let candidate = sorted.find(item => !coveredIds.has(item.itemId) && !selectedAnchorIds.has(item.itemId));

      if (!candidate) {
        candidate = selectFallbackAxisCandidate(sorted, selectedAnchorIds, axes);
      }

      if (!candidate) break;

      const axis = buildAxis(candidate, i, vehicles[i]);
      axes.push(axis);
      selectedAnchorIds.add(candidate.itemId);
      coveredIds = assignStagewiseCoveredIds(sorted, axes, threshold);
    }
    return axes;
  }

  function pushPendingMembers(people, axes, threshold) {
    const anchorIds = new Set(axes.map(axis => axis.anchorItemId));
    const unresolved = [];

    people.forEach(item => {
      if (anchorIds.has(item.itemId)) return;
      const matches = getPrimaryCandidateAxes(item, axes, threshold);
      if (!matches.length) {
        unresolved.push(item);
        return;
      }
      if (matches.length === 1) {
        matches[0].axis.pending.push(item);
        return;
      }
      unresolved.push(item);
    });

    axes.forEach(axis => {
      axis.pending.sort(byDistanceDesc);
      const freeSeats = Math.max(0, axis.capacity - axis.members.length);
      axis.members.push(...axis.pending.slice(0, freeSeats));
      axis.overflowed.push(...axis.pending.slice(freeSeats));
      axis.pending = [];
    });

    return unresolved.concat(axes.flatMap(axis => axis.overflowed));
  }

  function getAllAxisChoices(item, axes) {
    return axes
      .map(axis => ({ axis, diff: angleDiff(item.angleFromOrigin, axis.anchorAngle) }))
      .sort((a, b) => {
        if (a.diff !== b.diff) return a.diff - b.diff;
        if (b.axis.anchorDistanceKm !== a.axis.anchorDistanceKm) return b.axis.anchorDistanceKm - a.axis.anchorDistanceKm;
        return String(a.axis.axisId).localeCompare(String(b.axis.axisId), 'ja');
      });
  }

  function resolveUnsettledMembers(unresolved, axes) {
    const overflow = [];
    const sorted = [...unresolved].sort(byDistanceDesc);

    sorted.forEach(item => {
      const choices = getAllAxisChoices(item, axes);
      const target = choices.find(choice => choice.axis.members.length < choice.axis.capacity);
      if (!target) {
        overflow.push(item);
        return;
      }
      target.axis.members.push(item);
    });

    return overflow;
  }

  function finalizeAxisOrder(axes) {
    axes.forEach(axis => {
      axis.members.sort(byDistanceAsc);
    });
  }

  function buildAssignments(axes) {
    const assignments = [];
    axes.forEach(axis => {
      axis.members.forEach((item, index) => {
        assignments.push({
          item_id: item.itemId,
          actual_hour: item.actualHour,
          vehicle_id: axis.vehicleId,
          driver_name: axis.driverName,
          distance_km: item.distanceKm,
          stop_order: index + 1
        });
      });
    });
    return assignments;
  }

  function buildOverflowMeta(overflow) {
    return {
      overflowGroups: [],
      overflowEvaluations: [],
      capacityOverflowCount: overflow.length,
      capacityOverflowItems: overflow.map(item => ({
        itemId: item.itemId,
        hour: item.actualHour,
        group: item.area || '無し',
        distanceKm: item.distanceKm
      }))
    };
  }



  function getAxisRepresentative(axis, origin) {
    const members = Array.isArray(axis?.members) ? axis.members.filter(Boolean) : [];
    const sorted = [...members].sort(byDistanceDesc);
    const farthest = sorted[0] || null;
    return {
      axis,
      farthest,
      point: farthest?.point || null,
      distanceKm: toNumber(farthest?.distanceKm, 0),
      angleFromOrigin: Number.isFinite(farthest?.angleFromOrigin) ? farthest.angleFromOrigin : null,
      size: members.length
    };
  }


  function getMaxVehicleCapacity(vehicles) {
    return Math.max(0, ...(Array.isArray(vehicles) ? vehicles : []).map(vehicle => Number(vehicle?.capacity || 0)));
  }

  function buildAxisMergeCandidate(leftRep, rightRep, maxCapacity) {
    if (!leftRep?.axis || !rightRep?.axis) return null;
    if (!Number.isFinite(leftRep.angleFromOrigin) || !Number.isFinite(rightRep.angleFromOrigin)) return null;
    const combinedSize = Number(leftRep.size || 0) + Number(rightRep.size || 0);
    if (combinedSize <= 0 || combinedSize > maxCapacity) return null;

    const angleGap = angleDiff(leftRep.angleFromOrigin, rightRep.angleFromOrigin);
    if (!Number.isFinite(angleGap) || angleGap > SAME_DIRECTION_MERGE_ANGLE_THRESHOLD) return null;

    const pointDistanceKm = (leftRep.point && rightRep.point)
      ? haversineKm(leftRep.point, rightRep.point)
      : Infinity;
    if (Number.isFinite(pointDistanceKm) && pointDistanceKm > SAME_DIRECTION_MERGE_POINT_DISTANCE_KM) return null;

    const primary = Number(leftRep.distanceKm || 0) >= Number(rightRep.distanceKm || 0) ? leftRep : rightRep;
    const secondary = primary === leftRep ? rightRep : leftRep;

    return {
      primaryAxisId: primary.axis.axisId,
      secondaryAxisId: secondary.axis.axisId,
      angleGap,
      pointDistanceKm,
      combinedSize,
      primaryDistanceKm: Number(primary.distanceKm || 0),
      secondaryDistanceKm: Number(secondary.distanceKm || 0)
    };
  }

  function pickSameDirectionMergePair(axes, vehicles, origin) {
    if (!Array.isArray(axes) || axes.length < 2) return null;
    const maxCapacity = getMaxVehicleCapacity(vehicles);
    if (maxCapacity <= 0) return null;

    const reps = axes.map(axis => getAxisRepresentative(axis, origin));
    const candidates = [];

    for (let i = 0; i < reps.length; i += 1) {
      for (let j = i + 1; j < reps.length; j += 1) {
        const candidate = buildAxisMergeCandidate(reps[i], reps[j], maxCapacity);
        if (candidate) candidates.push(candidate);
      }
    }

    candidates.sort((a, b) => {
      if (a.angleGap !== b.angleGap) return a.angleGap - b.angleGap;
      if (a.pointDistanceKm !== b.pointDistanceKm) return a.pointDistanceKm - b.pointDistanceKm;
      if (b.primaryDistanceKm !== a.primaryDistanceKm) return b.primaryDistanceKm - a.primaryDistanceKm;
      if (a.combinedSize !== b.combinedSize) return a.combinedSize - b.combinedSize;
      return String(a.primaryAxisId).localeCompare(String(b.primaryAxisId), 'ja');
    });

    return candidates[0] || null;
  }

  function mergeFixedAxesByCandidate(axes, candidate) {
    if (!candidate?.primaryAxisId || !candidate?.secondaryAxisId) return Array.isArray(axes) ? [...axes] : [];
    const primary = axes.find(axis => axis.axisId === candidate.primaryAxisId);
    const secondary = axes.find(axis => axis.axisId === candidate.secondaryAxisId);
    if (!primary || !secondary) return [...axes];

    primary.members = [...(Array.isArray(primary.members) ? primary.members : []), ...(Array.isArray(secondary.members) ? secondary.members : [])];
    primary.overflowed = [...(Array.isArray(primary.overflowed) ? primary.overflowed : []), ...(Array.isArray(secondary.overflowed) ? secondary.overflowed : [])];
    primary.pending = [];
    return axes.filter(axis => axis.axisId !== candidate.secondaryAxisId);
  }

  function collapseSameDirectionAxes(axes, vehicles, origin) {
    let current = Array.isArray(axes) ? [...axes] : [];
    const standbyCount = Math.max(0, (Array.isArray(vehicles) ? vehicles.length : 0) - current.length);
    if (current.length < 2) return { axes: current, mergeLogs: [], standbyCount };

    const mergeLogs = [];
    while (current.length >= 2) {
      const candidate = pickSameDirectionMergePair(current, vehicles, origin);
      if (!candidate) break;
      mergeLogs.push(candidate);
      current = mergeFixedAxesByCandidate(current, candidate);
    }

    current.forEach(axis => {
      axis.members.sort(byDistanceAsc);
      const rep = getAxisRepresentative(axis, origin);
      if (rep?.farthest) {
        axis.anchorItemId = rep.farthest.itemId;
        axis.anchorName = rep.farthest.name;
        axis.anchorAngle = rep.farthest.angleFromOrigin;
        axis.anchorDistanceKm = rep.farthest.distanceKm;
      }
    });

    return {
      axes: current,
      mergeLogs,
      standbyCount: Math.max(0, (Array.isArray(vehicles) ? vehicles.length : 0) - current.length)
    };
  }

  function vehicleCanTakeAxis(vehicle, axisRep) {
    if (!vehicle || !axisRep) return false;
    if (Number(vehicle.capacity || 0) <= 0) return false;
    const need = Math.max(1, Number(axisRep.size || 0));
    return Number(vehicle.capacity || 0) >= need;
  }

  function pickEligibleVehiclesByScore(vehicles, axisReps) {
    const needed = axisReps.length;
    const remaining = [...vehicles];
    const selected = [];
    const requirements = [...axisReps].sort((a, b) => (b.size - a.size) || (b.distanceKm - a.distanceKm));

    requirements.forEach(axisRep => {
      if (selected.length >= needed) return;
      const eligible = remaining
        .filter(vehicle => vehicleCanTakeAxis(vehicle, axisRep))
        .sort((a, b) => {
          if (a.score !== b.score) return a.score - b.score;
          if (a.avgDistance !== b.avgDistance) return a.avgDistance - b.avgDistance;
          return String(a.driverName).localeCompare(String(b.driverName), 'ja');
        });
      const picked = eligible[0] || remaining.sort((a, b) => {
        if (a.score !== b.score) return a.score - b.score;
        return String(a.driverName).localeCompare(String(b.driverName), 'ja');
      })[0];
      if (!picked) return;
      selected.push(picked);
      const idx = remaining.findIndex(v => v.vehicleId === picked.vehicleId);
      if (idx >= 0) remaining.splice(idx, 1);
    });

    return selected;
  }

  function chooseNormalVehicleAssignments(axisReps, vehicles) {
    const selected = pickEligibleVehiclesByScore(vehicles, axisReps);
    const selectedMap = new Map(selected.map(v => [v.vehicleId, v]));
    const leftovers = vehicles.filter(v => !selectedMap.has(v.vehicleId));
    const pool = [...selected, ...leftovers];
    const assignments = new Map();
    const used = new Set();

    [...axisReps].sort((a, b) => (b.distanceKm - a.distanceKm) || (b.size - a.size)).forEach(axisRep => {
      const candidates = pool
        .filter(vehicle => !used.has(vehicle.vehicleId) && vehicleCanTakeAxis(vehicle, axisRep))
        .sort((a, b) => {
          if (a.score !== b.score) return a.score - b.score;
          if (a.avgDistance !== b.avgDistance) return a.avgDistance - b.avgDistance;
          return String(a.driverName).localeCompare(String(b.driverName), 'ja');
        });
      const picked = candidates[0] || pool.filter(vehicle => !used.has(vehicle.vehicleId)).sort((a, b) => {
        if (a.score !== b.score) return a.score - b.score;
        return String(a.driverName).localeCompare(String(b.driverName), 'ja');
      })[0];
      if (!picked) return;
      assignments.set(axisRep.axis.axisId, picked);
      used.add(picked.vehicleId);
    });

    return assignments;
  }

  function buildLastTripCandidate(axisRep, vehicle, origin) {
    if (!axisRep?.point || !vehicle?.homePoint) return null;
    if (!vehicleCanTakeAxis(vehicle, axisRep)) return null;
    const axisAngle = axisRep.angleFromOrigin;
    if (!Number.isFinite(axisAngle)) return null;
    const homeAngle = calcAngleFromOrigin(origin, vehicle.homePoint);
    const diff = angleDiff(axisAngle, homeAngle);
    return {
      axisId: axisRep.axis.axisId,
      vehicleId: vehicle.vehicleId,
      diff,
      homeDistanceKm: haversineKm(axisRep.point, vehicle.homePoint),
      axisDistanceKm: axisRep.distanceKm,
      size: axisRep.size
    };
  }

  function assignVehiclesToFixedAxes(fixedAxes, normalizedVehicles, origin) {
    const axisReps = fixedAxes.map(axis => getAxisRepresentative(axis, origin));
    const lastTripVehicles = normalizedVehicles.filter(vehicle => vehicle.isLastTrip);
    const normalVehicles = normalizedVehicles.filter(vehicle => !vehicle.isLastTrip);
    const axisMap = new Map(axisReps.map(rep => [rep.axis.axisId, rep]));
    const assigned = new Map();
    const usedVehicles = new Set();
    const usedAxes = new Set();

    if (!axisReps.length) return assigned;

    if (!lastTripVehicles.length) {
      return chooseNormalVehicleAssignments(axisReps, normalizedVehicles);
    }

    const pairCandidates = [];
    axisReps.forEach(axisRep => {
      lastTripVehicles.forEach(vehicle => {
        const entry = buildLastTripCandidate(axisRep, vehicle, origin);
        if (!entry || entry.diff > 90) return;
        pairCandidates.push(entry);
      });
    });

    pairCandidates.sort((a, b) => {
      if (a.homeDistanceKm !== b.homeDistanceKm) return a.homeDistanceKm - b.homeDistanceKm;
      if (a.diff !== b.diff) return a.diff - b.diff;
      if (b.axisDistanceKm !== a.axisDistanceKm) return b.axisDistanceKm - a.axisDistanceKm;
      return a.vehicleId - b.vehicleId;
    });

    pairCandidates.forEach(entry => {
      if (usedVehicles.has(entry.vehicleId) || usedAxes.has(entry.axisId)) return;
      const vehicle = normalizedVehicles.find(v => v.vehicleId === entry.vehicleId);
      if (!vehicle) return;
      assigned.set(entry.axisId, vehicle);
      usedVehicles.add(entry.vehicleId);
      usedAxes.add(entry.axisId);
    });

    const remainingLastTrip = lastTripVehicles.filter(vehicle => !usedVehicles.has(vehicle.vehicleId));

    remainingLastTrip.forEach(vehicle => {
      const choices = axisReps
        .filter(axisRep => !usedAxes.has(axisRep.axis.axisId))
        .map(axisRep => buildLastTripCandidate(axisRep, vehicle, origin))
        .filter(Boolean)
        .filter(entry => entry.diff <= 90)
        .sort((a, b) => {
          if (a.homeDistanceKm !== b.homeDistanceKm) return a.homeDistanceKm - b.homeDistanceKm;
          if (a.diff !== b.diff) return a.diff - b.diff;
          if (b.axisDistanceKm !== a.axisDistanceKm) return b.axisDistanceKm - a.axisDistanceKm;
          return a.vehicleId - b.vehicleId;
        });
      const picked = choices[0];
      if (!picked) return;
      assigned.set(picked.axisId, vehicle);
      usedVehicles.add(vehicle.vehicleId);
      usedAxes.add(picked.axisId);
    });

    let unassignedAxes = axisReps.filter(axisRep => !usedAxes.has(axisRep.axis.axisId));
    let freeNormalVehicles = normalVehicles.filter(vehicle => !usedVehicles.has(vehicle.vehicleId));
    let freeLastTripVehicles = lastTripVehicles.filter(vehicle => !usedVehicles.has(vehicle.vehicleId));

    if (unassignedAxes.length > 0 && freeLastTripVehicles.length > 0 && freeNormalVehicles.length < unassignedAxes.length) {
      freeLastTripVehicles.forEach(vehicle => {
        if (!unassignedAxes.length) return;
        if (freeNormalVehicles.length >= unassignedAxes.length) return;
        const choices = unassignedAxes
          .map(axisRep => buildLastTripCandidate(axisRep, vehicle, origin) || {
            axisId: axisRep.axis.axisId,
            vehicleId: vehicle.vehicleId,
            diff: Infinity,
            homeDistanceKm: vehicle.homePoint && axisRep.point ? haversineKm(axisRep.point, vehicle.homePoint) : Infinity,
            axisDistanceKm: axisRep.distanceKm,
            size: axisRep.size
          })
          .sort((a, b) => {
            if (a.homeDistanceKm !== b.homeDistanceKm) return a.homeDistanceKm - b.homeDistanceKm;
            if (a.diff !== b.diff) return a.diff - b.diff;
            if (b.axisDistanceKm !== a.axisDistanceKm) return b.axisDistanceKm - a.axisDistanceKm;
            return a.vehicleId - b.vehicleId;
          });
        const picked = choices[0];
        if (!picked) return;
        assigned.set(picked.axisId, vehicle);
        usedVehicles.add(vehicle.vehicleId);
        usedAxes.add(picked.axisId);
        unassignedAxes = axisReps.filter(axisRep => !usedAxes.has(axisRep.axis.axisId));
        freeNormalVehicles = normalVehicles.filter(v => !usedVehicles.has(v.vehicleId));
      });
    }

    unassignedAxes = axisReps.filter(axisRep => !usedAxes.has(axisRep.axis.axisId));
    const remainingVehicles = normalizedVehicles.filter(vehicle => !usedVehicles.has(vehicle.vehicleId));
    const normalAssignments = chooseNormalVehicleAssignments(unassignedAxes, remainingVehicles);
    normalAssignments.forEach((vehicle, axisId) => {
      assigned.set(axisId, vehicle);
      usedVehicles.add(vehicle.vehicleId);
      usedAxes.add(axisId);
    });

    return assigned;
  }

  function applyVehicleAssignmentsToAxes(fixedAxes, vehicleAssignments) {
    fixedAxes.forEach(axis => {
      const vehicle = vehicleAssignments.get(axis.axisId);
      if (!vehicle) return;
      axis.vehicleId = vehicle.vehicleId;
      axis.driverName = vehicle.driverName;
      axis.capacity = vehicle.capacity;
    });
  }

  function optimizeAssignments(items, vehicles, monthlyMap, options = {}) {
    const origin = getOrigin();
    const normalizedVehicles = (Array.isArray(vehicles) ? vehicles : [])
      .map(vehicle => normalizeVehicle(vehicle, monthlyMap))
      .filter(vehicle => vehicle.vehicleId > 0 && vehicle.capacity > 0);

    const normalizedItems = (Array.isArray(items) ? items : [])
      .map(item => normalizeItem(item, origin))
      .filter(item => item.itemId > 0 && Number.isFinite(item.distanceKm) && Number.isFinite(item.angleFromOrigin));

    if (!normalizedVehicles.length || !normalizedItems.length) {
      global.__THEMIS_LAST_OVERFLOW__ = buildOverflowMeta([]);
      return [];
    }

    const vehicleCount = Math.min(normalizedVehicles.length, normalizedItems.length);
    const threshold = Number(options?.axisThreshold || AREA_MEMBER_ANGLE_THRESHOLD);

    const fixedAxes = selectFixedAxes(normalizedItems, normalizedVehicles.slice(0, vehicleCount), threshold);
    try {
      debugLog('[DispatchCore][AXES]', fixedAxes.map(axis => ({
        axisId: axis.axisId,
        vehicleId: axis.vehicleId,
        driverName: axis.driverName,
        anchorName: axis.anchorName,
        anchorDistanceKm: Number(axis.anchorDistanceKm || 0),
        anchorAngle: Number(axis.anchorAngle || 0)
      })));
    } catch (error) {
      debugWarn('[DispatchCore][AXES] log failed:', error);
    }

    const unsettled = pushPendingMembers(normalizedItems, fixedAxes, threshold);
    const overflow = resolveUnsettledMembers(unsettled, fixedAxes);
    finalizeAxisOrder(fixedAxes);

    const mergeResult = collapseSameDirectionAxes(fixedAxes, normalizedVehicles.slice(0, vehicleCount), origin);
    const mergedAxes = Array.isArray(mergeResult?.axes) ? mergeResult.axes : fixedAxes;

    const vehicleAssignments = assignVehiclesToFixedAxes(mergedAxes, normalizedVehicles.slice(0, vehicleCount), origin);
    applyVehicleAssignmentsToAxes(mergedAxes, vehicleAssignments);

    const assignments = buildAssignments(mergedAxes);
    try {
      debugLog('[DispatchCore][RESULT]', mergedAxes.map(axis => ({
        axisId: axis.axisId,
        vehicleId: axis.vehicleId,
        driverName: axis.driverName,
        anchorName: axis.anchorName,
        memberNames: axis.members.map(item => item.name),
        memberCount: axis.members.length,
        capacity: axis.capacity
      })));
    } catch (error) {
      debugWarn('[DispatchCore][RESULT] log failed:', error);
    }
    global.__THEMIS_LAST_OVERFLOW__ = {
      ...buildOverflowMeta(overflow),
      fixedAxes: mergedAxes.map(axis => ({
        axisId: axis.axisId,
        vehicleId: axis.vehicleId,
        driverName: axis.driverName,
        anchorItemId: axis.anchorItemId,
        anchorName: axis.anchorName,
        anchorAngle: axis.anchorAngle,
        anchorDistanceKm: axis.anchorDistanceKm,
        memberItemIds: axis.members.map(item => item.itemId)
      })),
      totalSeatCapacity: mergedAxes.reduce((sum, axis) => sum + Number(axis.capacity || 0), 0),
      totalCastCount: normalizedItems.length,
      standbyVehicleCount: Number(mergeResult?.standbyCount || 0),
      mergedSameDirectionAxes: Array.isArray(mergeResult?.mergeLogs) ? mergeResult.mergeLogs : [],
      version: VERSION
    };

    return assignments;
  }

  function runDispatchPlan(origin, vehicles, people) {
    return optimizeAssignments(people, vehicles, null, {});
  }

  global.DispatchCore = {
    VERSION,
    optimizeAssignments,
    runDispatchPlan
  };
})(window);
