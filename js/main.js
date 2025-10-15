// New computeRTWStats implementation - replace the existing function body with this
function computeRTWStats(files, rows) {
    const sicknessCols = ['Sickness End','Sickness End Date','Sick End Date','SicknessEnd','Sick End','Sickness_End','End','End Date'];
    const dutyCols = ['Duty Date','DutyDate','Duty','Shift Date','ShiftDate','Next Duty','NextDuty','Next Duty Date','NextDutyDate'];
    const rtwCols = ['Return to Work','RTW','Return to work','Return','ReturnToWork','Return To Work Interview Completed','Return To Work Interview','RTW Interview Completed','RTW Interview','RTW Date','Return to Work Date'];
    const staffKeyFn = (r) => getStaffKey(r) || getValueByCanonical(r, 'Assignment No') || getValueByCanonical(r, 'Staff');

    const staffMap = {};
    const feedRows = [];

    // Build feedRows with rowIndex so we can link flags -> ranges by source
    if (Array.isArray(files) && files.length) {
        files.forEach(f => {
            (f.dataRows || []).forEach((r, idx) => {
                feedRows.push({ data: r, __locale: f._locale || 'uk', __fileId: f.id, __fileName: f.name, __rowIndex: idx });
            });
        });
    } else if (Array.isArray(rows) && rows.length) {
        // rows are wrappers (as returned by buildResults) — preserve any __rowIndex if present
        rows.forEach(r => feedRows.push({ ...r, __rowIndex: (r.__rowIndex != null ? r.__rowIndex : null) }));
    }

    // collect sickness ranges, duties and RTW flag objects (flags include parsedDate & source)
    const sicknessStartCols = ['Sickness Start','Sickness Start Date','Sick Start','Start','Start Date','SicknessStart'];

    feedRows.forEach(r => {
        const staff = String(staffKeyFn(r.data) || '').trim();
        if (!staff) return;
        staffMap[staff] = staffMap[staff] || { staff, sicknessRanges: [], dutyDates: [], rtwFlags: [] };
        const entry = staffMap[staff];

        // collect sickness start/end (prefer explicit start+end columns). Keep source reference.
        let startRaw = null, endRaw = null;
        for (const sc of sicknessCols) { const raw = getValueByCanonical(r.data, sc); if (raw != null) { endRaw = raw; break; } }
        for (const ss of sicknessStartCols) { const raw = getValueByCanonical(r.data, ss); if (raw != null) { startRaw = raw; break; } }
        try {
            const startDt = startRaw != null ? parseToDate(startRaw, r.__locale || 'uk') : null;
            const endDt = endRaw != null ? parseToDate(endRaw, r.__locale || 'uk') : null;
            const reasonCols = ['Reason','Sickness Reason','Absence Reason','Notes','Description'];
            let reasonVal = null;
            for (const rc of reasonCols) {
                const rv = getValueByCanonical(r.data, rc);
                if (rv != null && String(rv || '').trim() !== '') { reasonVal = String(rv).trim(); break; }
            }
            const source = { fileId: r.__fileId, fileName: r.__fileName, rowIndex: r.__rowIndex };

            if (endDt && !startDt) entry.sicknessRanges.push({ start: endDt, end: endDt, reason: reasonVal, sources: [source] });
            else if (startDt && !endDt) entry.sicknessRanges.push({ start: startDt, end: startDt, reason: reasonVal, sources: [source] });
            else if (startDt && endDt) entry.sicknessRanges.push({ start: startDt, end: endDt, reason: reasonVal, sources: [source] });
        } catch (e) {
            for (const sc of sicknessCols) {
                const raw = getValueByCanonical(r.data, sc);
                if (raw != null) {
                    const dt = parseToDate(raw, r.__locale || 'uk');
                    if (dt) entry.sicknessRanges.push({ start: dt, end: dt, reason: null, sources: [{ fileId: r.__fileId, fileName: r.__fileName, rowIndex: r.__rowIndex }] });
                    break;
                }
            }
        }

        // collect duty dates (skip Rest rows)
        for (const dc of dutyCols) {
            const raw = getValueByCanonical(r.data, dc);
            if (raw != null) {
                let dutyTypeVal = getValueByCanonical(r.data, 'Shift Type');
                if (dutyTypeVal == null) {
                    const tryAlts = ['ShiftType','Type','Roster Type','Assignment Info','Assignment'];
                    for (const alt of tryAlts) { const v = getValueByCanonical(r.data, alt); if (v != null) { dutyTypeVal = v; break; } }
                }
                let prio = -1;
                if (dutyTypeVal != null) {
                    const low = String(dutyTypeVal).trim().toLowerCase();
                    const tokens = low.split(/[^a-z]+/).filter(Boolean);
                    if (tokens.includes('combined')) prio = 3;
                    else if (tokens.includes('day') || tokens.includes('night')) prio = 2;
                    else if (tokens.includes('rest')) prio = 0;
                }
                if (prio < 0) {
                    try {
                        for (const key of Object.keys(r.data || {})) {
                            const s = (r.data[key] == null ? '' : (typeof r.data[key] === 'string' ? r.data[key] : JSON.stringify(r.data[key])));
                            if (!s) continue;
                            if (/\brest\b/i.test(s)) { prio = 0; break; }
                        }
                    } catch (e) { /* ignore */ }
                }
                if (prio < 0) break;
                if (prio === 0) break;
                const dt = parseToDate(raw, r.__locale || 'uk');
                if (dt) {
                    const iso = toLocalIso(dt);
                    entry.dutyDates.push({ date: dt, iso, prio, file: r.__fileName });
                }
                break;
            }
        }

        // collect RTW flags: store as objects with parsedDate (if available) + source
        for (const rc of rtwCols) {
            const raw = getValueByCanonical(r.data, rc);
            if (raw != null) {
                // attempt to find an associated date for the RTW flag: RTW Date column, or this row's sickness end, or duty date
                let dateRaw = getValueByCanonical(r.data, 'RTW Date') || getValueByCanonical(r.data, 'Return to Work Date') || endRaw;
                if (!dateRaw) {
                    for (const dc of dutyCols) {
                        const rd = getValueByCanonical(r.data, dc);
                        if (rd != null) { dateRaw = rd; break; }
                    }
                }
                const parsedDate = dateRaw ? parseToDate(dateRaw, r.__locale || 'uk') : null;
                entry.rtwFlags.push({
                    value: String(raw || '').trim(),
                    parsedDate: parsedDate,
                    iso: parsedDate ? toLocalIso(parsedDate) : null,
                    source: { fileId: r.__fileId, fileName: r.__fileName, rowIndex: r.__rowIndex }
                });
                break;
            }
        }
    });

    // Evaluate per-staff
    const results = [];
    const today = new Date(); today.setHours(0,0,0,0);

    Object.values(staffMap).forEach(s => {
        const requireReasonMatch = (typeof localStorage !== 'undefined' && localStorage.getItem('TERRA_MERGE_SICKNESS_BY_REASON') === 'true');
        const rawRanges = (s.sicknessRanges || []).filter(rr => rr && rr.start && rr.end).map(rr => ({
            start: new Date(rr.start.getFullYear(), rr.start.getMonth(), rr.start.getDate()),
            end: new Date(rr.end.getFullYear(), rr.end.getMonth(), rr.end.getDate()),
            reason: rr.reason != null ? String(rr.reason).trim() : null,
            sources: rr.sources || []
        }));
        rawRanges.sort((a,b) => a.start.getTime() - b.start.getTime());

        // merge adjacent/overlapping ranges while preserving sources
        const merged = [];
        for (const r0 of rawRanges) {
            if (!merged.length) { merged.push({ start: new Date(r0.start), end: new Date(r0.end), count:1, reason: r0.reason || null, sources: (r0.sources || []).slice() }); continue; }
            const last = merged[merged.length - 1];
            const nextDayAfterLast = new Date(last.end.getTime()); nextDayAfterLast.setDate(nextDayAfterLast.getDate() + 1);
            const lastReason = last.reason != null ? String(last.reason).trim() : null;
            const thisReason = r0.reason != null ? String(r0.reason).trim() : null;
            if (r0.start.getTime() <= nextDayAfterLast.getTime() && (!requireReasonMatch || lastReason === thisReason)) {
                if (r0.end.getTime() > last.end.getTime()) last.end = new Date(r0.end);
                last.count = (last.count || 1) + 1;
                last.sources = (last.sources || []).concat(r0.sources || []);
            } else {
                merged.push({ start: new Date(r0.start), end: new Date(r0.end), count: 1, reason: r0.reason || null, sources: (r0.sources || []).slice() });
            }
        }

        // latest merged range considered the 'current' sickness episode
        const latestRange = merged.length ? merged[merged.length - 1] : null;
        const sickness = latestRange ? new Date(latestRange.end.getTime()) : null;

        // process dutyDates (keep highest priority per date)
        let processedDuties = [];
        if (s.dutyDates.length) {
            const byIso = {};
            s.dutyDates.forEach(dd => {
                const iso = dd.iso || (dd instanceof Date ? toLocalIso(dd) : null);
                if (!iso) return;
                if (!byIso[iso] || (dd.prio != null && dd.prio > byIso[iso].prio)) byIso[iso] = dd;
            });
            processedDuties = Object.values(byIso).map(x => x.date).sort((a,b) => a.getTime() - b.getTime());
        }

        // nextDuty and lastDuty as before
        const nextDuty = processedDuties.find(d => {
            const d0 = new Date(d.getFullYear(), d.getMonth(), d.getDate());
            return d0.getTime() >= today.getTime();
        }) || null;
        const lastOnOrBefore = processedDuties.filter(d => {
            const d0 = new Date(d.getFullYear(), d.getMonth(), d.getDate());
            return d0.getTime() <= today.getTime();
        });
        const lastDuty = lastOnOrBefore.length ? lastOnOrBefore[lastOnOrBefore.length - 1] : null;

        // helper to identify yes-like RTW flag values
        const yesMatches = (f) => {
            if (!f || !f.value) return false;
            const v = String(f.value || '').toLowerCase();
            return v === 'y' || v === 'yes' || v === 'true' || v === '1';
        };

        // determine if an RTW flag belongs to the given merged range
        const flagBelongsToRange = (flag, range) => {
            if (!flag) return false;
            // Prefer date association when present. Allow a short post-return window.
            if (flag.parsedDate instanceof Date && !isNaN(flag.parsedDate)) {
                const fd = new Date(flag.parsedDate.getFullYear(), flag.parsedDate.getMonth(), flag.parsedDate.getDate());
                const POST_WINDOW_DAYS = 14; // configurable window after sickness end to accept RTW recorded shortly after return
                const windowEnd = new Date(range.end.getTime() + (POST_WINDOW_DAYS * 24*60*60*1000));
                return fd.getTime() >= range.start.getTime() && fd.getTime() <= windowEnd.getTime();
            }
            // If no usable date, accept the flag if it comes from a source row that contributed to this range
            try {
                if (flag.source && Array.isArray(range.sources) && range.sources.length) {
                    return range.sources.some(src => (src && src.fileId === flag.source.fileId && (src.rowIndex == null || flag.source.rowIndex == null || src.rowIndex === flag.source.rowIndex)));
                }
            } catch (e) { /* ignore */ }
            // Conservative fallback: do not attribute the flag to the range
            return false;
        };

        // compute rtwDone only considering flags that belong to the latestRange
        let rtwDoneForLatest = false;
        if (latestRange) {
            for (const flag of (s.rtwFlags || [])) {
                try {
                    if (!yesMatches(flag)) continue;
                    if (flagBelongsToRange(flag, latestRange)) { rtwDoneForLatest = true; break; }
                } catch (e) { /* ignore per-flag errors */ }
            }
        } else {
            // fallback: no sickness range detected — use broad heuristic (any yes-like RTW flag)
            rtwDoneForLatest = (s.rtwFlags || []).some(f => yesMatches(f));
        }

        // remaining semantics preserved
        const sicknessEnded = !!(sickness && sickness.getTime() <= today.getTime());
        const hadShiftAfter = sickness && lastDuty && lastDuty.getTime() >= sickness.getTime();
        const hadShiftAfterSickness = !!hadShiftAfter;
        const rtwInterviewDone = !!rtwDoneForLatest;
        const rawMultiple = Array.isArray(rawRanges) && rawRanges.length > 1;
        const continuingSickness = !!(latestRange && (latestRange.end.getTime() >= today.getTime() || (latestRange.count && latestRange.count > 1))) || rawMultiple;
        const mergedRangesOut = merged.map(mr => ({ start: toLocalIso(mr.start), end: toLocalIso(mr.end), parts: mr.count || 1, reason: (mr.reason != null ? String(mr.reason) : null) }));

        results.push({
            staff: s.staff,
            sickness,
            sicknessEnded,
            hadShiftAfterSickness,
            hadShiftAfter: !!hadShiftAfter,
            rtwInterviewDone,
            rtwDone: !!rtwDoneForLatest,
            sampleDuty: processedDuties.length ? processedDuties[0] : null,
            nextDuty,
            lastDuty,
            continuingSickness,
            sicknessRanges: mergedRangesOut
        });
    });

    return results;
}
