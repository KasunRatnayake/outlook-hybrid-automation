import { chromium, type Page } from '@playwright/test';

/**
 * WebView2 DevTools often answers on `localhost` only; `127.0.0.1` can time out while
 * `http://localhost:9222/json/list` works (same host, different loopback resolution on Windows).
 */
const OUTLOOK_CDP_URL = process.env.OUTLOOK_CDP_URL ?? 'http://localhost:9222';

/** Set `POC_SKIP_CDP=1` to finish after GIPHY opens — Outlook WebView2 usually does not expose Chrome DevTools on 9222 unless you configure it. */
const POC_SKIP_CDP = process.env.POC_SKIP_CDP === '1';

/**
 * Comma-separated ports or full URLs, e.g. `9222,9229` — Office/VS Code samples sometimes use 9229.
 * Overrides single OUTLOOK_CDP_URL when set.
 */
function resolveCdpUrlsToTry(): string[] {
    const portsEnv = process.env.OUTLOOK_CDP_PORTS?.trim();
    if (portsEnv) {
        return portsEnv.split(',').map((s) => {
            const t = s.trim();
            if (t.startsWith('http://') || t.startsWith('https://')) {
                return t;
            }
            return `http://localhost:${t}`;
        });
    }
    if (process.env.OUTLOOK_CDP_URL !== undefined) {
        return [OUTLOOK_CDP_URL];
    }
    return [
        'http://localhost:9222',
        'http://127.0.0.1:9222',
        'http://localhost:9229',
        'http://127.0.0.1:9229'
    ];
}

/** Chromium exposes JSON metadata here when remote debugging is enabled. */
async function probeCdpHttpRoot(rootUrl: string): Promise<boolean> {
    const base = rootUrl.replace(/\/$/, '');
    for (const path of ['/json/list', '/json/version', '/json']) {
        try {
            const r = await fetch(`${base}${path}`, {
                signal: AbortSignal.timeout(2500)
            });
            if (r.ok) {
                return true;
            }
        } catch {
            /* ignore */
        }
    }
    return false;
}

function logWebView2CdpInstructions(): void {
    console.log(`
 Why ETIMEDOUT on 127.0.0.1:9222 while http://localhost:9222/json/list works
 ───────────────────────────────────────────────────────────────────────────
 The debugger may only bind to "localhost" (IPv4/IPv6), not 127.0.0.1. This POC
 defaults to http://localhost:9222. Override with OUTLOOK_CDP_URL if needed.

 If nothing responds: Embedded WebView2 does NOT open a CDP port unless the browser
 process was started with --remote-debugging-port=… .

 Enable it (pick one), then fully quit Outlook and start it again:

 1) User env var (affects WebView2 created after login / new processes):
    WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9222

 2) Registry (per Microsoft Learn — WebView2 policy):
    HKCU\\Software\\Policies\\Microsoft\\Edge\\WebView2\\AdditionalBrowserArguments
    String value: --remote-debugging-port=9222

 After Outlook restarts, open GIPHY again and run this script. Try OUTLOOK_CDP_PORTS=9222,9229
 if 9222 is busy. Use POC_SKIP_CDP=1 to skip Playwright until CDP works.

 Docs: https://learn.microsoft.com/en-us/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-chromium
       https://learn.microsoft.com/en-us/microsoft-edge/webview2/how-to/debug-visual-studio-code
`);
}

async function connectCdpWithRetry(
    url: string,
    attempts = 8,
    delayMs = 1500
): Promise<Awaited<ReturnType<typeof chromium.connectOverCDP>>> {
    let lastErr: unknown;
    for (let i = 0; i < attempts; i++) {
        try {
            return await chromium.connectOverCDP(url);
        } catch (e) {
            lastErr = e;
            if (i < attempts - 1) {
                await new Promise((r) => setTimeout(r, delayMs));
            }
        }
    }
    throw new Error(
        `CDP connect failed after ${String(attempts)} tries (${url}). ` +
            `Last error: ${lastErr instanceof Error ? lastErr.message : String(lastErr)}`
    );
}

async function connectPlaywrightToOutlookWebViewCdp(): Promise<
    Awaited<ReturnType<typeof chromium.connectOverCDP>>
> {
    const urls = resolveCdpUrlsToTry();
    let anyHttp = false;
    for (const u of urls) {
        if (await probeCdpHttpRoot(u)) {
            anyHttp = true;
            console.log(` CDP HTTP probe OK (${u}) — DevTools endpoint is up.`);
            break;
        }
    }
    if (!anyHttp) {
        console.log(
            ' CDP probe: no response on /json/version — WebView2 is almost certainly running without --remote-debugging-port.'
        );
        logWebView2CdpInstructions();
    }

    const errors: string[] = [];
    for (const url of urls) {
        try {
            console.log(` Trying Playwright connectOverCDP → ${url}`);
            return await connectCdpWithRetry(url, 6, 1200);
        } catch (e) {
            errors.push(`${url}: ${e instanceof Error ? e.message : String(e)}`);
        }
    }
    throw new Error(
        `CDP failed for all URLs (${urls.join(', ')}). ` +
            (anyHttp
                ? 'HTTP responded but Playwright could not attach — try closing other Edge/WebView2 debug sessions.'
                : 'Enable WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS (see log above). ') +
            `Details: ${errors.join(' | ')}`
    );
}

/**
 * GIPHY grid images are often inside links; first `img` may be a logo. Try several strategies.
 */
async function clickFirstGiphySearchResult(page: Page): Promise<void> {
    await page.waitForLoadState('domcontentloaded');
    await new Promise((r) => setTimeout(r, 600));

    const tryRun = async (label: string, fn: () => Promise<void>): Promise<boolean> => {
        try {
            console.log(` GIPHY: trying click — ${label}`);
            await fn();
            return true;
        } catch {
            return false;
        }
    };

    if (
        await tryRun('link-wrapped thumbnail (common grid)', async () => {
            const loc = page
                .locator(
                    'a[href*="giphy.com"] img, a[href*="/gifs/"] img, [role="listbox"] img, [class*="Grid"] img'
                )
                .first();
            await loc.waitFor({ state: 'visible', timeout: 15_000 });
            await loc.click({ force: true, timeout: 10_000 });
        })
    ) {
        return;
    }

    if (
        await tryRun('img[src*="media" or giphy CDN]', async () => {
            const loc = page
                .locator('img[src*="giphy"], img[src*="media0.giphy.com"], img[src*="media"]')
                .first();
            await loc.waitFor({ state: 'visible', timeout: 12_000 });
            await loc.click({ force: true });
        })
    ) {
        return;
    }

    if (
        await tryRun('center click on first large visible img (mouse)', async () => {
            const imgs = page.locator('img');
            const n = await imgs.count();
            if (n === 0) {
                throw new Error('no img elements');
            }
            for (let i = 0; i < Math.min(n, 30); i++) {
                const el = imgs.nth(i);
                if (!(await el.isVisible())) {
                    continue;
                }
                const box = await el.boundingBox();
                if (box && box.width >= 48 && box.height >= 48) {
                    await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
                    return;
                }
            }
            const box = await imgs.nth(0).boundingBox();
            if (!box) {
                throw new Error('no bounding box');
            }
            await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
        })
    ) {
        return;
    }

    if (
        await tryRun('HTMLElement.click() in page (bypass overlay hit-testing)', async () => {
            const clicked = await page.evaluate(() => {
                const imgs = Array.from(document.querySelectorAll('img'));
                const pick =
                    imgs.find((img) => {
                        const r = img.getBoundingClientRect();
                        return r.width >= 64 && r.height >= 64 && r.top >= 0 && r.bottom <= window.innerHeight + 200;
                    }) ?? imgs.find((img) => img.getBoundingClientRect().width > 32);
                if (pick) {
                    pick.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
                    (pick as HTMLElement).click();
                    return true;
                }
                return false;
            });
            if (!clicked) {
                throw new Error('evaluate found no suitable img');
            }
        })
    ) {
        return;
    }

    throw new Error(
        'Could not click a GIF result — try narrowing selectors for your GIPHY UI version.'
    );
}

async function runHybridPOC() {
    console.log(' Starting WinAppDriver session (Outlook will launch)...');
    const baseUrl = 'http://127.0.0.1:4723';
    const outlookExe =
        'C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE';

    // app: 'Root' only attaches to the desktop — it does NOT start Outlook.
    // Use the Outlook EXE so WinAppDriver launches the application.
    const sessionResponse = await fetch(`${baseUrl}/session`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            desiredCapabilities: {
                platformName: 'Windows',
                deviceName: 'WindowsPC',
                app: outlookExe
            }
        })
    });

    const sessionData = (await sessionResponse.json()) as {
        sessionId?: string;
        value?: { message?: string };
    };
    if (!sessionResponse.ok || !sessionData.sessionId) {
        throw new Error(
            `Session failed: ${sessionData.value?.message ?? sessionResponse.statusText}`
        );
    }
    let sessionId = sessionData.sessionId;
    /** When true, `replaceSessionWithRoot` is active — safe to find compose without deleting a main-Outlook HWND session first. */
    let sessionUsesDesktopRoot = false;

    await fetch(`${baseUrl}/session/${sessionId}/timeouts`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ implicit: 100 })
    });

    async function click(selector: string, isName = false) {
        const strategy = isName ? "name" : "accessibility id";
        console.log(`🔎 Looking for: ${selector} via ${strategy}`);

        const res = await fetch(`${baseUrl}/session/${sessionId}/element`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ using: strategy, value: selector })
        });
        const data = await res.json() as any;
        const eid = data.value?.['element-6066-11e4-a52e-4f735466cecf'] ?? data.value?.ELEMENT;

        if (!eid) throw new Error(`Could not find ${selector}`);

        await fetch(`${baseUrl}/session/${sessionId}/element/${eid}/click`, { method: 'POST' });
    }

    /** Window title is usually "Inbox - … - Outlook", not the literal name "Outlook", so use xpath. */
    async function clickXpath(xpath: string, label: string) {
        console.log(`🔎 Looking for: ${label}`);
        const res = await fetch(`${baseUrl}/session/${sessionId}/element`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ using: 'xpath', value: xpath })
        });
        const data = await res.json() as any;
        const eid = data.value?.['element-6066-11e4-a52e-4f735466cecf'] ?? data.value?.ELEMENT;
        if (!eid) throw new Error(`Could not find ${label}`);
        await fetch(`${baseUrl}/session/${sessionId}/element/${eid}/click`, { method: 'POST' });
    }

    async function tryFindXpath(xpath: string): Promise<string | null> {
        const res = await fetch(`${baseUrl}/session/${sessionId}/element`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ using: 'xpath', value: xpath })
        });
        const data = await res.json() as any;
        const eid = data.value?.['element-6066-11e4-a52e-4f735466cecf'] ?? data.value?.ELEMENT;
        if (!res.ok || !eid) {
            return null;
        }
        return eid;
    }

    /** All matches (e.g. several compose windows); WinAppDriver `element` returns only the first. */
    async function findAllElementIdsByXpath(xpath: string): Promise<string[]> {
        const res = await fetch(`${baseUrl}/session/${sessionId}/elements`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ using: 'xpath', value: xpath })
        });
        const data = (await res.json()) as {
            value?: Array<Record<string, string>>;
        };
        if (!res.ok || !Array.isArray(data.value)) {
            return [];
        }
        const ids: string[] = [];
        for (const el of data.value) {
            const id =
                el['element-6066-11e4-a52e-4f735466cecf'] ?? (el as { ELEMENT?: string }).ELEMENT;
            if (id) {
                ids.push(String(id));
            }
        }
        return ids;
    }

    const WIN_ELEMENT_ID = 'element-6066-11e4-a52e-4f735466cecf';

    async function getElementSize(
        eid: string
    ): Promise<{ width: number; height: number } | null> {
        const r = await fetch(`${baseUrl}/session/${sessionId}/element/${eid}/size`, {
            method: 'GET'
        });
        const d = (await r.json()) as { value?: { width?: number; height?: number } };
        if (!r.ok || d.value == null) {
            return null;
        }
        const w = d.value.width ?? 0;
        const h = d.value.height ?? 0;
        if (w <= 0 || h <= 0) {
            return null;
        }
        return { width: w, height: h };
    }

    async function getElementLocation(eid: string): Promise<{ x: number; y: number } | null> {
        const r = await fetch(`${baseUrl}/session/${sessionId}/element/${eid}/location`, {
            method: 'GET'
        });
        const d = (await r.json()) as { value?: { x?: number; y?: number } };
        if (!r.ok || d.value == null) {
            return null;
        }
        return { x: d.value.x ?? 0, y: d.value.y ?? 0 };
    }

    /**
     * Default `/click` can hit the wrong spot on ribbon buttons; use pointer at bbox center.
     * Tries element-relative move first, then viewport coordinates from location+size.
     * @param yFraction vertical hit position in the control (0 = top, 1 = bottom). Ribbon icons often sit ~upper-mid (~0.38–0.45), not geometric 0.5.
     */
    async function pointerClickCenterOfElement(
        eid: string,
        options?: { yFraction?: number }
    ): Promise<boolean> {
        const yFraction = options?.yFraction ?? 0.5;
        const sz = await getElementSize(eid);
        if (!sz) {
            return false;
        }
        const ox = Math.floor(sz.width / 2);
        const oy = Math.max(1, Math.floor(sz.height * yFraction) - 1);
        const originEl = { [WIN_ELEMENT_ID]: eid };

        let r = await fetch(`${baseUrl}/session/${sessionId}/actions`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                actions: [
                    {
                        type: 'pointer',
                        id: 'giphy-ptr',
                        parameters: { pointerType: 'mouse' },
                        actions: [
                            {
                                type: 'pointerMove',
                                duration: 80,
                                origin: originEl,
                                x: ox,
                                y: oy
                            },
                            { type: 'pointerDown', button: 0 },
                            { type: 'pointerUp', button: 0 }
                        ]
                    }
                ]
            })
        });
        await fetch(`${baseUrl}/session/${sessionId}/actions`, { method: 'DELETE' }).catch(() => {});
        if (r.ok) {
            console.log(
                ` Pointer click at (${String(ox)}, ${String(oy)}) in control (yFraction=${String(yFraction)})`
            );
            return true;
        }

        const loc = await getElementLocation(eid);
        if (!loc) {
            return false;
        }
        const cx = Math.round(loc.x + sz.width / 2);
        const cy = Math.round(loc.y + oy);
        r = await fetch(`${baseUrl}/session/${sessionId}/actions`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                actions: [
                    {
                        type: 'pointer',
                        id: 'giphy-ptr2',
                        parameters: { pointerType: 'mouse' },
                        actions: [
                            {
                                type: 'pointerMove',
                                duration: 80,
                                origin: 'viewport',
                                x: cx,
                                y: cy
                            },
                            { type: 'pointerDown', button: 0 },
                            { type: 'pointerUp', button: 0 }
                        ]
                    }
                ]
            })
        });
        await fetch(`${baseUrl}/session/${sessionId}/actions`, { method: 'DELETE' }).catch(() => {});
        if (r.ok) {
            console.log(` Clicked center via viewport (~${String(cx)}, ~${String(cy)})`);
            return true;
        }
        return false;
    }

    async function findElementByStrategy(
        strategy: 'name' | 'accessibility id',
        value: string
    ): Promise<string | null> {
        const res = await fetch(`${baseUrl}/session/${sessionId}/element`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ using: strategy, value })
        });
        const data = await res.json() as { value?: Record<string, string> };
        const eid = data.value?.[WIN_ELEMENT_ID] ?? (data.value as { ELEMENT?: string })?.ELEMENT;
        if (!res.ok || !eid) {
            return null;
        }
        return String(eid);
    }

    /**
     * Accessibility Insights tree: … → NUIDocumentWindow → … → Button 'New Email' (click the Button, not Text 50020).
     */
    const NEW_EMAIL_BUTTON_XPATHS: { xpath: string; label: string }[] = [
        {
            xpath: `//Pane[@Name='NUIDocumentWindow']/descendant::Button[@Name='New Email']`,
            label: 'NUIDocumentWindow → Button[New Email]'
        },
        {
            xpath: `//Window[contains(@Name,'Inbox') and contains(@Name,'Outlook')]/descendant::Button[@Name='New Email']`,
            label: 'Inbox+Outlook window → Button[New Email]'
        },
        {
            xpath: `//Window[contains(@Name,'Outlook')]/descendant::Button[@Name='New Email']`,
            label: 'Outlook window → Button[New Email]'
        },
        { xpath: `//Button[@Name='New Email']`, label: '//Button[@Name=New Email]' },
        {
            xpath: `//*[@ControlType='Button' and @Name='New Email']`,
            label: 'ControlType Button + Name New Email'
        },
        { xpath: `//Button[contains(@Name,'New Email')]`, label: 'Button name contains New Email' },
        {
            xpath: `//Custom/descendant::Button[@Name='New Email']`,
            label: 'Custom → Button[New Email] (Insights path)'
        },
        {
            xpath: `//*[@ControlType='50000' and @Name='New Email']`,
            label: 'ControlType 50000 (UIA Button) + Name'
        }
    ];

    /**
     * Main Outlook explorer frame only — not the separate compose window ("… Message (HTML)").
     * Avoid `contains(@Name,'@')` alone: it can match another app's top-level window before
     * the inbox HWND is ready, so WinAppDriver binds to the wrong tree and "New Email" vanishes.
     */
    const OUTLOOK_WINDOW_XPATHS = [
        `//Window[contains(@Name,'Outlook') and not(contains(@Name,'Message (HTML)'))]`,
        `//Window[contains(@Name,'Microsoft Outlook')]`,
        `//Window[contains(@Name,'outlook')]`,
        `//Window[contains(@Name,'@') and contains(@Name,'Outlook')]`,
        `//Window[contains(@Name,' - Mail')]`,
        `//Window[contains(@Name,'Inbox')]`
    ];

    /**
     * Compose is a separate top-level window from the main Outlook frame; Insights:
     * window "Untitled - Message (HTML)" → … → Group "GIPHY" → Button "GIPHY".
     */
    const COMPOSE_WINDOW_XPATHS = [
        `//Window[contains(@Name,'Message (HTML)')]`,
        `//Window[contains(@Name,'Untitled') and contains(@Name,'Message')]`,
        `//Window[contains(@Name,'Message') and contains(@Name,'HTML')]`,
        `//Window[contains(@Name,'(HTML)')]`,
        `//Window[contains(@Name,'- Message') and not(contains(@Name,'Outlook'))]`
    ];

    /**
     * Button-only — never match Group[@Name='GIPHY'] (large box; name lookup hits Group first).
     */
    const GIPHY_BUTTON_XPATHS: { xpath: string; label: string }[] = [
        {
            xpath: `//Group[@Name='Message']//Group[@Name='GIPHY']/Button[@Name='GIPHY']`,
            label: 'Insights: Message → Group GIPHY → Button'
        },
        {
            xpath: `//Group[@Name='GIPHY']/Button[@Name='GIPHY']`,
            label: 'Group GIPHY → Button'
        },
        {
            xpath: `//Group[@Name='Message']//Button[@Name='GIPHY']`,
            label: 'Message group → Button GIPHY'
        },
        { xpath: `//Pane[contains(@Name,'Ribbon')]//Button[@Name='GIPHY']`, label: 'Ribbon → Button GIPHY' },
        { xpath: `//Button[@Name='GIPHY']`, label: '//Button[@Name=GIPHY]' },
        {
            xpath: `//*[@ControlType='50000' and @Name='GIPHY']`,
            label: 'ControlType 50000 + Name GIPHY'
        }
    ];

    /** Last resort — may match Group; try only after Button xpaths fail. */
    const GIPHY_LOOSE_XPATHS: { xpath: string; label: string }[] = [
        {
            xpath: `//Button[contains(@Name,'GIPHY')]`,
            label: 'Button name contains GIPHY'
        },
        { xpath: `//*[contains(@Name,'GIPHY')]`, label: 'Any Name contains GIPHY (may be Group)' }
    ];

    /** WebDriver Control key for chord with "n". */
    const WD_CTRL = '\uE009';

    async function getWindowHandles(): Promise<string[]> {
        for (const path of ['window/handles', 'window_handles'] as const) {
            const r = await fetch(`${baseUrl}/session/${sessionId}/${path}`, { method: 'GET' });
            if (r.ok) {
                const d = (await r.json()) as { value?: string[] };
                if (Array.isArray(d.value) && d.value.length > 0) {
                    return d.value;
                }
            }
        }
        return [];
    }

    async function switchToWindow(handle: string): Promise<void> {
        const r = await fetch(`${baseUrl}/session/${sessionId}/window`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name: handle })
        });
        if (!r.ok) {
            throw new Error(await r.text());
        }
    }

    /**
     * /keys runs against the session's "current" window; HWND sessions often need an explicit window switch
     * or a click on the mail surface first — otherwise: no such window / selected window closed.
     */
    async function prepareKeyboardFocus(): Promise<void> {
        const handles = await getWindowHandles();
        for (const h of handles) {
            try {
                await switchToWindow(h);
                await new Promise((r) => setTimeout(r, 300));
                console.log(' Switched WinAppDriver to a window handle before keys');
                return;
            } catch {
                /* try next */
            }
        }
        for (const xp of [`//Pane[@Name='NUIDocumentWindow']`, ...OUTLOOK_WINDOW_XPATHS]) {
            const id = await tryFindXpath(xp);
            if (id) {
                await fetch(`${baseUrl}/session/${sessionId}/element/${id}/click`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({})
                });
                await new Promise((r) => setTimeout(r, 500));
                console.log(' Clicked focus target before keys:', xp.slice(0, 50));
                return;
            }
        }
    }

    async function sendCtrlNviaW3CActions(): Promise<void> {
        const r = await fetch(`${baseUrl}/session/${sessionId}/actions`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                actions: [
                    {
                        type: 'key',
                        id: 'kbd',
                        actions: [
                            { type: 'keyDown', value: WD_CTRL },
                            { type: 'keyDown', value: 'n' },
                            { type: 'keyUp', value: 'n' },
                            { type: 'keyUp', value: WD_CTRL }
                        ]
                    }
                ]
            })
        });
        await fetch(`${baseUrl}/session/${sessionId}/actions`, { method: 'DELETE' }).catch(() => {});
        if (!r.ok) {
            throw new Error(await r.text());
        }
    }

    async function sendCtrlN(): Promise<void> {
        await prepareKeyboardFocus();

        let r = await fetch(`${baseUrl}/session/${sessionId}/keys`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ value: [WD_CTRL, 'n'] })
        });
        if (r.ok) {
            console.log(' Sent Ctrl+N via /keys');
            return;
        }
        const errText = await r.text();
        console.log(' /keys failed, trying W3C actions…', errText.slice(0, 200));
        try {
            await sendCtrlNviaW3CActions();
            console.log(' Sent Ctrl+N via W3C actions');
        } catch (e2) {
            const a = e2 instanceof Error ? e2.message : String(e2);
            throw new Error(`Ctrl+N failed. /keys: ${errText}; actions: ${a}`);
        }
    }

    async function findOutlookWindowElementId(
        maxRounds: number,
        pauseMs: number
    ): Promise<string | null> {
        for (let round = 0; round < maxRounds; round++) {
            for (const xp of OUTLOOK_WINDOW_XPATHS) {
                const id = await tryFindXpath(xp);
                if (id) {
                    return id;
                }
            }
            await new Promise((r) => setTimeout(r, pauseMs));
        }
        return null;
    }

    async function findComposeWindowElementId(
        maxRounds: number,
        pauseMs: number
    ): Promise<string | null> {
        for (let round = 0; round < maxRounds; round++) {
            for (const xp of COMPOSE_WINDOW_XPATHS) {
                const all = await findAllElementIdsByXpath(xp);
                if (all.length > 1) {
                    console.log(
                        ` Found ${String(all.length)} compose candidate(s); using last (newest/top heuristic): ${xp.slice(0, 56)}…`
                    );
                }
                if (all.length > 0) {
                    return all[all.length - 1]!;
                }
                const one = await tryFindXpath(xp);
                if (one) {
                    return one;
                }
            }
            await new Promise((r) => setTimeout(r, pauseMs));
        }
        return null;
    }

    /** Ribbon groups (e.g. GIPHY) may not be in the UIA tree until the Message strip is active. */
    async function tryFocusComposeRibbon(): Promise<void> {
        const targets = [
            `//Pane[contains(@Name,'Lower Ribbon')]`,
            `//Pane[@Name='Lower Ribbon']`,
            `//Pane[contains(@Name,'Ribbon')]`,
            `//TabItem[@Name='Message']`,
            `//TabItem[contains(@Name,'Message')]`
        ];
        for (const xp of targets) {
            const id = await tryFindXpath(xp);
            if (!id) {
                continue;
            }
            await fetch(`${baseUrl}/session/${sessionId}/element/${id}/click`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({})
            });
            await new Promise((r) => setTimeout(r, 500));
            console.log(' Activated compose ribbon context:', xp.slice(0, 58));
            return;
        }
    }

    async function bindSessionToOutlookWindowElement(winId: string): Promise<boolean> {
        const attrRes = await fetch(
            `${baseUrl}/session/${sessionId}/element/${winId}/attribute/NativeWindowHandle`,
            { method: 'GET' }
        );
        const aj = (await attrRes.json()) as { value?: string | number | null };
        const v = aj.value;
        if (v == null || v === '') {
            console.log('⚠️ No NativeWindowHandle on window element');
            return false;
        }
        const hwndNum = typeof v === 'number' ? v : parseInt(String(v), 10);
        if (Number.isNaN(hwndNum) || hwndNum === 0) {
            return false;
        }
        const hex = '0x' + hwndNum.toString(16);
        const r2 = await fetch(`${baseUrl}/session`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                desiredCapabilities: {
                    platformName: 'Windows',
                    deviceName: 'WindowsPC',
                    appTopLevelWindow: hex
                }
            })
        });
        const d2 = (await r2.json()) as { sessionId?: string; value?: { message?: string } };
        if (!r2.ok || !d2.sessionId) {
            console.log('⚠️ appTopLevelWindow session failed:', d2.value?.message);
            return false;
        }
        try {
            await fetch(`${baseUrl}/session/${sessionId}`, { method: 'DELETE' });
        } catch {
            /* ignore */
        }
        sessionId = d2.sessionId;
        sessionUsesDesktopRoot = false;
        await fetch(`${baseUrl}/session/${sessionId}/timeouts`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ implicit: 100 })
        });
        console.log(' Reattached session to window HWND', hex);
        return true;
    }

    /**
     * EXE session often does not expose the main Outlook window in the UIA tree (splash / wrong scope).
     */
    async function reattachToOutlookHwnd(): Promise<boolean> {
        await new Promise((r) => setTimeout(r, 6000));
        const winId = await findOutlookWindowElementId(40, 1000);
        if (!winId) {
            console.log('⚠️ HWND reattach skipped (Outlook window not in EXE session tree)');
            return false;
        }
        return bindSessionToOutlookWindowElement(winId);
    }

    /** Replace current session with Desktop Root (sees all top-level HWNDs). */
    async function replaceSessionWithRoot(): Promise<void> {
        try {
            await fetch(`${baseUrl}/session/${sessionId}`, { method: 'DELETE' });
        } catch {
            /* ignore */
        }
        const r = await fetch(`${baseUrl}/session`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                desiredCapabilities: {
                    platformName: 'Windows',
                    deviceName: 'WindowsPC',
                    app: 'Root'
                }
            })
        });
        const d = (await r.json()) as { sessionId?: string; value?: { message?: string } };
        if (!r.ok || !d.sessionId) {
            throw new Error(`Root session failed: ${d.value?.message ?? r.statusText}`);
        }
        sessionId = d.sessionId;
        sessionUsesDesktopRoot = true;
        await fetch(`${baseUrl}/session/${sessionId}/timeouts`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ implicit: 100 })
        });
    }

    /**
     * Desktop Root only — do NOT bind to main Outlook HWND here.
     * Binding main then deleting that session after compose opens was destabilizing the new-mail window.
     */
    async function recoverSessionViaRoot(): Promise<void> {
        console.log(' Trying WinAppDriver Root session to find Outlook on the desktop…');
        await replaceSessionWithRoot();
        await new Promise((r) => setTimeout(r, 1500));

        const winId = await findOutlookWindowElementId(50, 1500);
        if (!winId) {
            throw new Error(
                'Outlook main window not found from Root. Is Outlook running and visible?'
            );
        }
        console.log(
            ' Recovery: staying on Desktop Root (no main-window HWND bind) so compose is not affected when switching to the message window.'
        );
    }

    /**
     * GIPHY lives on the compose window only. An HWND session on the main frame does not include it.
     * Root → find "Untitled - Message (HTML)" (or similar) → appTopLevelWindow on that HWND.
     */
    async function reattachSessionToComposeWindow(): Promise<void> {
        console.log(' Scoping automation to the compose (message) window for GIPHY…');
        if (!sessionUsesDesktopRoot) {
            console.log(
                ' Switching to Desktop Root to locate the compose window (avoids ending a main-Outlook HWND session after new mail is open).'
            );
            await replaceSessionWithRoot();
        } else {
            console.log(' Already on Desktop Root — skipping extra session reset before compose.');
        }
        await new Promise((r) => setTimeout(r, 600));

        const winId = await findComposeWindowElementId(50, 1000);
        if (!winId) {
            throw new Error(
                'Compose window not found (e.g. Untitled - Message (HTML)). Is it open and visible?'
            );
        }
        const ok = await bindSessionToOutlookWindowElement(winId);
        if (!ok) {
            throw new Error('Could not bind WinAppDriver to the compose window HWND');
        }
        console.log(' Session is scoped to the compose (Message) window.');
        await new Promise((r) => setTimeout(r, 1800));
    }

    /** After a ribbon click, the add-in search box is a strong signal. Avoid //*[Name*=GIPHY] — it matches the ribbon button that was already there. */
    async function waitForGiphyPaneHint(timeoutMs = 10_000): Promise<boolean> {
        const hints = [
            `//Edit[contains(@Name,'Search')]`,
            `//ComboBox[contains(@Name,'Search')]`,
            `//Spinner[contains(@Name,'Search')]`,
            `//Document[contains(translate(@Name,'GIPHY','giphy'),'giphy')]`,
            `//Pane[contains(translate(@Name,'GIPHY','giphy'),'giphy')]`
        ];
        const deadline = Date.now() + timeoutMs;
        while (Date.now() < deadline) {
            for (const xp of hints) {
                const id = await tryFindXpath(xp);
                if (id) {
                    return true;
                }
            }
            await new Promise((r) => setTimeout(r, 450));
        }
        return false;
    }

    async function clickGiphyButtonWithRetry(timeoutMs = 45_000): Promise<void> {
        await tryFocusComposeRibbon();
        const deadline = Date.now() + timeoutMs;
        let attempt = 0;

        /** Ribbon add-in buttons: icon is usually upper half of the UIA rect — avoid y=0.5 when bbox includes label. */
        const RIBBON_ICON_Y = 0.42;

        async function tryGiphyElement(eid: string, label: string): Promise<boolean> {
            console.log(`🔎 Targeting GIPHY control: ${label}`);
            let clicked = false;
            if (await pointerClickCenterOfElement(eid, { yFraction: RIBBON_ICON_Y })) {
                clicked = true;
            } else if (await pointerClickCenterOfElement(eid, { yFraction: 0.5 })) {
                clicked = true;
            } else {
                const clickRes = await fetch(`${baseUrl}/session/${sessionId}/element/${eid}/click`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({})
                });
                clicked = clickRes.ok;
            }
            if (!clicked) {
                return false;
            }
            const paneOk = await waitForGiphyPaneHint(10_000);
            if (!paneOk) {
                console.log(
                    '⚠️ Click sent but GIPHY/search UI never showed in UIA — likely missed the icon; retrying…'
                );
            } else {
                console.log(' GIPHY task pane / search UI visible in UIA.');
            }
            return paneOk;
        }

        while (Date.now() < deadline) {
            attempt += 1;
            if (attempt > 1 && attempt % 4 === 0) {
                await tryFocusComposeRibbon();
            }

            // Do NOT use findElementByStrategy('name','GIPHY') — it returns Group "GIPHY" before Button "GIPHY".

            for (const { xpath, label } of GIPHY_BUTTON_XPATHS) {
                const ids = await findAllElementIdsByXpath(xpath);
                if (ids.length > 0) {
                    for (let j = ids.length - 1; j >= 0; j--) {
                        const id = ids[j];
                        if (id && (await tryGiphyElement(id, `${label} [match ${String(j)}]`))) {
                            return;
                        }
                    }
                } else {
                    const found = await tryFindXpath(xpath);
                    if (found && (await tryGiphyElement(found, label))) {
                        return;
                    }
                }
            }

            const allNamedButtons = await findAllElementIdsByXpath(`//Button[@Name='GIPHY']`);
            if (allNamedButtons.length > 1) {
                console.log(
                    ` Found ${String(allNamedButtons.length)} Button[@Name=GIPHY]; trying last → first`
                );
            }
            for (let i = allNamedButtons.length - 1; i >= 0; i--) {
                const bid = allNamedButtons[i];
                if (bid && (await tryGiphyElement(bid, `//Button[@Name='GIPHY'] [${String(i)}]`))) {
                    return;
                }
            }

            for (const { xpath, label } of GIPHY_LOOSE_XPATHS) {
                const found = await tryFindXpath(xpath);
                if (!found) {
                    continue;
                }
                if (await tryGiphyElement(found, label)) {
                    return;
                }
            }

            const a11y = await findElementByStrategy('accessibility id', 'GIPHY');
            if (a11y && (await tryGiphyElement(a11y, 'accessibility id'))) {
                return;
            }

            await new Promise((r) => setTimeout(r, 1200));
        }
        throw new Error(
            'GIPHY ribbon click did not open the add-in (search UI never appeared in UIA) within timeout'
        );
    }

    async function clickNewEmailButtonWithRetry(timeoutMs = 45_000): Promise<void> {
        const deadline = Date.now() + timeoutMs;
        while (Date.now() < deadline) {
            for (const { xpath, label } of NEW_EMAIL_BUTTON_XPATHS) {
                const eid = await tryFindXpath(xpath);
                if (!eid) {
                    continue;
                }
                console.log(`🔎 Found New Email: ${label}`);
                const clickRes = await fetch(`${baseUrl}/session/${sessionId}/element/${eid}/click`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({})
                });
                if (clickRes.ok) {
                    return;
                }
            }
            await new Promise((r) => setTimeout(r, 1500));
        }
        console.log(' XPath exhausted; trying Ctrl+N…');
        await sendCtrlN();
        await new Promise((r) => setTimeout(r, 2000));
    }

    try {
        await new Promise((r) => setTimeout(r, 5000));
        const hwndBound = await reattachToOutlookHwnd();
        if (!hwndBound) {
            await recoverSessionViaRoot();
        } else {
            console.log(
                ' Releasing main Outlook HWND session; using Desktop Root for New Mail (prevents tearing that session down after compose opens).'
            );
            await replaceSessionWithRoot();
        }

        // Focus main Outlook window so name/xpath searches hit the right UI tree.
        let focused = false;
        for (const xp of OUTLOOK_WINDOW_XPATHS) {
            try {
                await clickXpath(xp, `main window (${xp.slice(0, 40)}…)`);
                console.log('Focused Outlook main window');
                focused = true;
                break;
            } catch {
                /* try next */
            }
        }
        if (!focused) {
            console.log('⚠️ Could not click Outlook window (continuing anyway)');
        }
        await new Promise((r) => setTimeout(r, 800));

        // STEP 1: New Email — Insights: Button Name "New Email" under NUIDocumentWindow (not the inner Text node).
        try {
            await click('New Email', true);
        } catch {
            try {
                await click('New mail', true);
            } catch {
                await clickNewEmailButtonWithRetry();
            }
        }
        console.log('New mail action done');

        await new Promise((r) => setTimeout(r, 2000));

        // STEP 2: GIPHY — Insights: Button "GIPHY" is under the compose window
        // "Untitled - Message (HTML)", not the main Outlook HWND session.
        await reattachSessionToComposeWindow();
        await clickGiphyButtonWithRetry();
        console.log(' GIPHY add-in UI detected.');

        if (POC_SKIP_CDP) {
            console.log(
                ' POC_SKIP_CDP=1 — skipping Playwright (WebView2 in Outlook rarely exposes CDP on 9222 without extra setup).'
            );
            console.log('🎉 SUCCESS (WinAppDriver + GIPHY side panel).');
        } else {
            // STEP 3: Playwright — WebView2 must be started with --remote-debugging-port (see connectPlaywrightToOutlookWebViewCdp).
            console.log('🔗 Connecting Playwright over CDP (WebView2 / Chromium DevTools)…');
            const browser = await connectPlaywrightToOutlookWebViewCdp();
            const page = browser.contexts()[0].pages().find((p) => p.url().includes('giphy'));

            if (!page) throw new Error('Giphy WebView not found. Is the add-in pane open?');

            await page.fill('input[placeholder*="Search"]', 'cat');
            await page.keyboard.press('Enter');
            await clickFirstGiphySearchResult(page);

            console.log('🎉 SUCCESS (full hybrid: desktop + Playwright).');
        }
    } catch (e) {
        console.error(' Failed:', e instanceof Error ? e.message : String(e));
    } finally {
        console.log(' Ending WinAppDriver session (automation disconnect — Outlook windows stay open).');
        await fetch(`${baseUrl}/session/${sessionId}`, { method: 'DELETE' });
    }
}

runHybridPOC();

