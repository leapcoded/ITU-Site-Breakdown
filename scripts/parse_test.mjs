// Lightweight test for parseDateByLocale and formatByLocale copied/adapted from js/ui.js
const localStorage = { store: {}, getItem(k){ return this.store[k]; }, setItem(k,v){ this.store[k]=String(v); } };

function formatToUK(val) {
    if (val == null) return '';
    if (val instanceof Date) {
        const d = val.getDate().toString().padStart(2,'0');
        const m = (val.getMonth()+1).toString().padStart(2,'0');
        const y = val.getFullYear();
        return `${d}/${m}/${y}`;
    }
    if (typeof val === 'string') {
        const iso = val.match(/^\s*(\d{4})-(\d{2})-(\d{2})(?:T.*)?\s*$/);
        if (iso) return `${iso[3]}/${iso[2]}/${iso[1]}`;
        const sl = val.match(/^\s*(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s*$/);
        if (sl) {
            const day = sl[1].padStart(2,'0');
            const mon = sl[2].padStart(2,'0');
            const year = sl[3].length === 2 ? (parseInt(sl[3],10) > 50 ? '19'+sl[3] : '20'+sl[3]) : sl[3];
            return `${day}/${mon}/${year}`;
        }
        const dash = val.match(/^\s*(\d{1,2})-(\d{1,2})-(\d{2,4})\s*$/);
        if (dash) {
            const day = dash[1].padStart(2,'0');
            const mon = dash[2].padStart(2,'0');
            const year = dash[3].length === 2 ? (parseInt(dash[3],10) > 50 ? '19'+dash[3] : '20'+dash[3]) : dash[3];
            return `${day}/${mon}/${year}`;
        }
        const m2 = val.match(/^\s*(\d{1,2})\s+([A-Za-z]{3,})\.?\s*,?\s*(\d{4})\s*$/);
        if (m2) {
            const day = m2[1].padStart(2,'0');
            const monStr = m2[2].toLowerCase();
            const months = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
            const mi = months.indexOf(monStr.slice(0,3));
            if (mi !== -1) return `${day}/${String(mi+1).padStart(2,'0')}/${m2[3]}`;
        }
    }
    return String(val);
}

function formatByLocale(val, locale) {
    const loc = locale || 'uk';
    if (val == null) return '';
    if (loc === 'us') {
        if (typeof val === 'string') {
            const iso = val.match(/^\s*(\d{4})-(\d{2})-(\d{2})(?:T.*)?\s*$/);
            if (iso) return `${iso[2]}/${iso[3]}/${iso[1]}`;
            const slash = val.match(/^\s*(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s*$/);
            if (slash) {
                const day = slash[1].padStart(2,'0');
                const mon = slash[2].padStart(2,'0');
                const y = slash[3].length === 2 ? (parseInt(slash[3],10) > 50 ? '19'+slash[3] : '20'+slash[3]) : slash[3];
                return `${mon}/${day}/${y}`;
            }
        }
        return String(val);
    }
    return formatToUK(val);
}

function parseDateByLocale(str, locale = 'uk') {
    if (!str || typeof str !== 'string') return null;
    const s = str.trim();
        const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        const toISO = (y, m, d) => `${String(y).padStart(4,'0')}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    const normYear = (y) => { if (y < 100) return y + (y > 50 ? 1900 : 2000); return y; };
    const valid = (y, m, d) => { const dt = new Date(y, m - 1, d); return dt.getFullYear() === y && dt.getMonth() === m - 1 && dt.getDate() === d; };

    if (isoMatch) {
        const y = parseInt(isoMatch[1], 10);
        const m = parseInt(isoMatch[2], 10);
        const d = parseInt(isoMatch[3], 10);
            if (valid(y, m, d)) return toISO(y, m, d);
    }

    const loc = (locale || 'uk').toString().toLowerCase();
    const userAmbig = localStorage.getItem('TERRA_DATE_AMBIGUOUS') || 'auto';

    const slashMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (slashMatch) {
        let a = parseInt(slashMatch[1], 10);
        let b = parseInt(slashMatch[2], 10);
        let y = normYear(parseInt(slashMatch[3], 10));

        const tryUk = () => { if (valid(y, b, a)) return toISO(y, b, a); return null; };
        const tryUs = () => { if (valid(y, a, b)) return toISO(y, a, b); return null; };

        if (loc === 'uk') {
            const r = tryUk();
            if (r) return r;
            return tryUs();
        }
        if (loc === 'us') {
            const r = tryUs();
            if (r) return r;
            return tryUk();
        }

        if (a > 12 && b <= 12) {
            return tryUk();
        }
        if (b > 12 && a <= 12) {
            return tryUs();
        }
        if (userAmbig === 'uk') {
            const rUk = tryUk(); if (rUk) return rUk; const rUs = tryUs(); if (rUs) return rUs; return null;
        }
        if (userAmbig === 'us') {
            const rUs = tryUs(); if (rUs) return rUs; const rUk = tryUk(); if (rUk) return rUk; return null;
        }
        const rUk = tryUk(); if (rUk) return rUk; const rUs = tryUs(); if (rUs) return rUs; return null;
    }

    const dayMonthName = s.match(/^\s*(\d{1,2})\s+([A-Za-z]{3,})\.?\s*,?\s*(\d{4})\s*$/);
    if (dayMonthName) {
        const day = parseInt(dayMonthName[1], 10);
        const monStr = dayMonthName[2].toLowerCase().slice(0,3);
        const months = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
        const mi = months.indexOf(monStr);
        if (mi !== -1) {
            const month = mi + 1;
            const year = parseInt(dayMonthName[3], 10);
                if (valid(year, month, day)) return toISO(year, month, day);
        }
    }
    const monthDayName = s.match(/^\s*([A-Za-z]{3,})\.?\,?\s*(\d{1,2}),?\s*(\d{4})\s*$/);
    if (monthDayName) {
        const monStr = monthDayName[1].toLowerCase().slice(0,3);
        const day = parseInt(monthDayName[2], 10);
        const months = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
        const mi = months.indexOf(monStr);
        if (mi !== -1) {
            const month = mi + 1;
            const year = parseInt(monthDayName[3], 10);
                if (valid(year, month, day)) return toISO(year, month, day);
        }
    }

    const parsed = Date.parse(s);
    if (!Number.isNaN(parsed)) {
        const dt = new Date(parsed);
            return toISO(dt.getFullYear(), dt.getMonth() + 1, dt.getDate());
    }
    return null;
}

// quick tests
function assertEq(actual, expected, name){
    if (actual === expected) console.log('PASS:', name);
    else console.error('FAIL:', name, 'expected', expected, 'got', actual);
}

// tests
assertEq(parseDateByLocale('12/27/23','us'),'2023-12-27','US short year');
assertEq(parseDateByLocale('27/12/2023','uk'),'2023-12-27','UK long year');
assertEq(parseDateByLocale('03/04/2025','us'),'2025-03-04','ambiguous as US');
assertEq(parseDateByLocale('03/04/2025','uk'),'2025-04-03','ambiguous as UK');

// format
assertEq(formatByLocale('2023-12-27','uk'),'27/12/2023','format uk');
assertEq(formatByLocale('2023-12-27','us'),'12/27/2023','format us');

console.log('done');
