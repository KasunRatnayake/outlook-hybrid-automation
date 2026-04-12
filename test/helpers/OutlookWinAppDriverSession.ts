import { Agent, fetch as undiciFetch } from 'undici';

/**
 * WinAppDriver: long POST /session may block until Outlook launches — extended undici timeouts.
 */
export class OutlookWinAppDriverSession {
    private static readonly WIN_ELEMENT_ID = 'element-6066-11e4-a52e-4f735466cecf';

    static readonly sessionDispatcher = new Agent({
        connectTimeout: 120_000,
        headersTimeout: 900_000,
        bodyTimeout: 900_000
    });

    private readonly baseUrl = 'http://127.0.0.1:4723';
    private sessionId: string;
    private sessionUsesDesktopRoot = false;

    private constructor(sessionId: string) {
        this.sessionId = sessionId;
    }

    static async start(): Promise<OutlookWinAppDriverSession> {
        console.log(' Starting WinAppDriver session (Outlook will launch)...');
        const baseUrl = 'http://127.0.0.1:4723';
        const outlookExe =
            'C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE';

        const sessionResponse = await undiciFetch(`${baseUrl}/session`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                desiredCapabilities: {
                    platformName: 'Windows',
                    deviceName: 'WindowsPC',
                    app: outlookExe
                }
            }),
            dispatcher: OutlookWinAppDriverSession.sessionDispatcher
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
        const inst = new OutlookWinAppDriverSession(sessionData.sessionId);
        await fetch(`${inst.baseUrl}/session/${inst.sessionId}/timeouts`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ implicit: 100 })
        });
        return inst;
    }

    async dispose(): Promise<void> {
        console.log(' Ending WinAppDriver session (automation disconnect — Outlook windows stay open).');
        await fetch(`${this.baseUrl}/session/${this.sessionId}`, { method: 'DELETE' });
    }

    async runDesktopFlowToGiphyPane(): Promise<void> {
        console.log(' Part 1 — WinAppDriver (Outlook → GIPHY pane)…');
        await new Promise((r) => setTimeout(r, 5000));
        const hwndBound = await this.reattachToOutlookHwnd();
        if (!hwndBound) {
            await this.recoverSessionViaRoot();
        } else {
            console.log(
                ' Releasing main Outlook HWND session; using Desktop Root for New Mail (prevents tearing that session down after compose opens).'
            );
            await this.replaceSessionWithRoot();
        }

        // Focus main Outlook window so name/xpath searches hit the right UI tree.
        let focused = false;
        for (const xp of OutlookWinAppDriverSession.OUTLOOK_WINDOW_XPATHS) {
            try {
                await this.clickXpath(xp, `main window (${xp.slice(0, 40)}…)`);
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
            await this.click('New Email', true);
        } catch {
            try {
                await this.click('New mail', true);
            } catch {
                await this.clickNewEmailButtonWithRetry();
            }
        }
        console.log('New mail action done');

        await new Promise((r) => setTimeout(r, 2000));

        // STEP 2: GIPHY — Insights: Button "GIPHY" is under the compose window
        // "Untitled - Message (HTML)", not the main Outlook HWND session.
        await this.reattachSessionToComposeWindow();
        await this.clickGiphyButtonWithRetry();
        console.log(' GIPHY add-in UI detected.');
    }

    private async click(selector: string, isName = false) {
        const strategy = isName ? "name" : "accessibility id";
        console.log(`🔎 Looking for: ${selector} via ${strategy}`);

        const res = await fetch(`${this.baseUrl}/session/${this.sessionId}/element`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ using: strategy, value: selector })
        });
        const data = await res.json() as any;
        const eid = data.value?.['element-6066-11e4-a52e-4f735466cecf'] ?? data.value?.ELEMENT;

        if (!eid) throw new Error(`Could not find ${selector}`);

        await fetch(`${this.baseUrl}/session/${this.sessionId}/element/${eid}/click`, { method: 'POST' });
    }

    /** Window title is usually "Inbox - … - Outlook", not the literal name "Outlook", so use xpath. */
    private async clickXpath(xpath: string, label: string) {
        console.log(`🔎 Looking for: ${label}`);
        const res = await fetch(`${this.baseUrl}/session/${this.sessionId}/element`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ using: 'xpath', value: xpath })
        });
        const data = await res.json() as any;
        const eid = data.value?.['element-6066-11e4-a52e-4f735466cecf'] ?? data.value?.ELEMENT;
        if (!eid) throw new Error(`Could not find ${label}`);
        await fetch(`${this.baseUrl}/session/${this.sessionId}/element/${eid}/click`, { method: 'POST' });
    }

    private async tryFindXpath(xpath: string): Promise<string | null> {
        const res = await fetch(`${this.baseUrl}/session/${this.sessionId}/element`, {
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
    private async findAllElementIdsByXpath(xpath: string): Promise<string[]> {
        const res = await fetch(`${this.baseUrl}/session/${this.sessionId}/elements`, {
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


    private async getElementSize(
        eid: string
    ): Promise<{ width: number; height: number } | null> {
        const r = await fetch(`${this.baseUrl}/session/${this.sessionId}/element/${eid}/size`, {
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

    private async getElementLocation(eid: string): Promise<{ x: number; y: number } | null> {
        const r = await fetch(`${this.baseUrl}/session/${this.sessionId}/element/${eid}/location`, {
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
    private async pointerClickCenterOfElement(
        eid: string,
        options?: { yFraction?: number }
    ): Promise<boolean> {
        const yFraction = options?.yFraction ?? 0.5;
        const sz = await this.getElementSize(eid);
        if (!sz) {
            return false;
        }
        const ox = Math.floor(sz.width / 2);
        const oy = Math.max(1, Math.floor(sz.height * yFraction) - 1);
        const originEl = { [OutlookWinAppDriverSession.WIN_ELEMENT_ID]: eid };

        let r = await fetch(`${this.baseUrl}/session/${this.sessionId}/actions`, {
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
        await fetch(`${this.baseUrl}/session/${this.sessionId}/actions`, { method: 'DELETE' }).catch(() => {});
        if (r.ok) {
            console.log(
                ` Pointer click at (${String(ox)}, ${String(oy)}) in control (yFraction=${String(yFraction)})`
            );
            return true;
        }

        const loc = await this.getElementLocation(eid);
        if (!loc) {
            return false;
        }
        const cx = Math.round(loc.x + sz.width / 2);
        const cy = Math.round(loc.y + oy);
        r = await fetch(`${this.baseUrl}/session/${this.sessionId}/actions`, {
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
        await fetch(`${this.baseUrl}/session/${this.sessionId}/actions`, { method: 'DELETE' }).catch(() => {});
        if (r.ok) {
            console.log(` Clicked center via viewport (~${String(cx)}, ~${String(cy)})`);
            return true;
        }
        return false;
    }

    private async findElementByStrategy(
        strategy: 'name' | 'accessibility id',
        value: string
    ): Promise<string | null> {
        const res = await fetch(`${this.baseUrl}/session/${this.sessionId}/element`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ using: strategy, value })
        });
        const data = await res.json() as { value?: Record<string, string> };
        const eid = data.value?.[OutlookWinAppDriverSession.WIN_ELEMENT_ID] ?? (data.value as { ELEMENT?: string })?.ELEMENT;
        if (!res.ok || !eid) {
            return null;
        }
        return String(eid);
    }

    /**
     * Accessibility Insights tree: … → NUIDocumentWindow → … → Button 'New Email' (click the Button, not Text 50020).
     */
    private static readonly NEW_EMAIL_BUTTON_XPATHS: { xpath: string; label: string }[] = [
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
    private static readonly OUTLOOK_WINDOW_XPATHS = [
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
    private static readonly COMPOSE_WINDOW_XPATHS = [
        `//Window[contains(@Name,'Message (HTML)')]`,
        `//Window[contains(@Name,'Untitled') and contains(@Name,'Message')]`,
        `//Window[contains(@Name,'Message') and contains(@Name,'HTML')]`,
        `//Window[contains(@Name,'(HTML)')]`,
        `//Window[contains(@Name,'- Message') and not(contains(@Name,'Outlook'))]`
    ];

    /**
     * Button-only — never match Group[@Name='GIPHY'] (large box; name lookup hits Group first).
     */
    private static readonly GIPHY_BUTTON_XPATHS: { xpath: string; label: string }[] = [
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
    private static readonly GIPHY_LOOSE_XPATHS: { xpath: string; label: string }[] = [
        {
            xpath: `//Button[contains(@Name,'GIPHY')]`,
            label: 'Button name contains GIPHY'
        },
        { xpath: `//*[contains(@Name,'GIPHY')]`, label: 'Any Name contains GIPHY (may be Group)' }
    ];

    /** WebDriver Control key for chord with "n". */
    private static readonly WD_CTRL = '\uE009';

    private async getWindowHandles(): Promise<string[]> {
        for (const path of ['window/handles', 'window_handles'] as const) {
            const r = await fetch(`${this.baseUrl}/session/${this.sessionId}/${path}`, { method: 'GET' });
            if (r.ok) {
                const d = (await r.json()) as { value?: string[] };
                if (Array.isArray(d.value) && d.value.length > 0) {
                    return d.value;
                }
            }
        }
        return [];
    }

    private async switchToWindow(handle: string): Promise<void> {
        const r = await fetch(`${this.baseUrl}/session/${this.sessionId}/window`, {
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
    private async prepareKeyboardFocus(): Promise<void> {
        const handles = await this.getWindowHandles();
        for (const h of handles) {
            try {
                await this.switchToWindow(h);
                await new Promise((r) => setTimeout(r, 300));
                console.log(' Switched WinAppDriver to a window handle before keys');
                return;
            } catch {
                /* try next */
            }
        }
        for (const xp of [`//Pane[@Name='NUIDocumentWindow']`, ...OutlookWinAppDriverSession.OUTLOOK_WINDOW_XPATHS]) {
            const id = await this.tryFindXpath(xp);
            if (id) {
                await fetch(`${this.baseUrl}/session/${this.sessionId}/element/${id}/click`, {
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

    private async sendCtrlNviaW3CActions(): Promise<void> {
        const r = await fetch(`${this.baseUrl}/session/${this.sessionId}/actions`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                actions: [
                    {
                        type: 'key',
                        id: 'kbd',
                        actions: [
                            { type: 'keyDown', value: OutlookWinAppDriverSession.WD_CTRL },
                            { type: 'keyDown', value: 'n' },
                            { type: 'keyUp', value: 'n' },
                            { type: 'keyUp', value: OutlookWinAppDriverSession.WD_CTRL }
                        ]
                    }
                ]
            })
        });
        await fetch(`${this.baseUrl}/session/${this.sessionId}/actions`, { method: 'DELETE' }).catch(() => {});
        if (!r.ok) {
            throw new Error(await r.text());
        }
    }

    private async sendCtrlN(): Promise<void> {
        await this.prepareKeyboardFocus();

        let r = await fetch(`${this.baseUrl}/session/${this.sessionId}/keys`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ value: [OutlookWinAppDriverSession.WD_CTRL, 'n'] })
        });
        if (r.ok) {
            console.log(' Sent Ctrl+N via /keys');
            return;
        }
        const errText = await r.text();
        console.log(' /keys failed, trying W3C actions…', errText.slice(0, 200));
        try {
            await this.sendCtrlNviaW3CActions();
            console.log(' Sent Ctrl+N via W3C actions');
        } catch (e2) {
            const a = e2 instanceof Error ? e2.message : String(e2);
            throw new Error(`Ctrl+N failed. /keys: ${errText}; actions: ${a}`);
        }
    }

    private async findOutlookWindowElementId(
        maxRounds: number,
        pauseMs: number
    ): Promise<string | null> {
        for (let round = 0; round < maxRounds; round++) {
            for (const xp of OutlookWinAppDriverSession.OUTLOOK_WINDOW_XPATHS) {
                const id = await this.tryFindXpath(xp);
                if (id) {
                    return id;
                }
            }
            await new Promise((r) => setTimeout(r, pauseMs));
        }
        return null;
    }

    private async findComposeWindowElementId(
        maxRounds: number,
        pauseMs: number
    ): Promise<string | null> {
        for (let round = 0; round < maxRounds; round++) {
            for (const xp of OutlookWinAppDriverSession.COMPOSE_WINDOW_XPATHS) {
                const all = await this.findAllElementIdsByXpath(xp);
                if (all.length > 1) {
                    console.log(
                        ` Found ${String(all.length)} compose candidate(s); using last (newest/top heuristic): ${xp.slice(0, 56)}…`
                    );
                }
                if (all.length > 0) {
                    return all[all.length - 1]!;
                }
                const one = await this.tryFindXpath(xp);
                if (one) {
                    return one;
                }
            }
            await new Promise((r) => setTimeout(r, pauseMs));
        }
        return null;
    }

    /** Ribbon groups (e.g. GIPHY) may not be in the UIA tree until the Message strip is active. */
    private async tryFocusComposeRibbon(): Promise<void> {
        const targets = [
            `//Pane[contains(@Name,'Lower Ribbon')]`,
            `//Pane[@Name='Lower Ribbon']`,
            `//Pane[contains(@Name,'Ribbon')]`,
            `//TabItem[@Name='Message']`,
            `//TabItem[contains(@Name,'Message')]`
        ];
        for (const xp of targets) {
            const id = await this.tryFindXpath(xp);
            if (!id) {
                continue;
            }
            await fetch(`${this.baseUrl}/session/${this.sessionId}/element/${id}/click`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({})
            });
            await new Promise((r) => setTimeout(r, 500));
            console.log(' Activated compose ribbon context:', xp.slice(0, 58));
            return;
        }
    }

    private async bindSessionToOutlookWindowElement(winId: string): Promise<boolean> {
        const attrRes = await fetch(
            `${this.baseUrl}/session/${this.sessionId}/element/${winId}/attribute/NativeWindowHandle`,
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
        const r2 = await undiciFetch(`${this.baseUrl}/session`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                desiredCapabilities: {
                    platformName: 'Windows',
                    deviceName: 'WindowsPC',
                    appTopLevelWindow: hex
                }
            }),
            dispatcher: OutlookWinAppDriverSession.sessionDispatcher
        });
        const d2 = (await r2.json()) as { sessionId?: string; value?: { message?: string } };
        if (!r2.ok || !d2.sessionId) {
            console.log('⚠️ appTopLevelWindow session failed:', d2.value?.message);
            return false;
        }
        try {
            await fetch(`${this.baseUrl}/session/${this.sessionId}`, { method: 'DELETE' });
        } catch {
            /* ignore */
        }
        this.sessionId = d2.sessionId;
        this.sessionUsesDesktopRoot = false;
        await fetch(`${this.baseUrl}/session/${this.sessionId}/timeouts`, {
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
    private async reattachToOutlookHwnd(): Promise<boolean> {
        await new Promise((r) => setTimeout(r, 6000));
        const winId = await this.findOutlookWindowElementId(40, 1000);
        if (!winId) {
            console.log('⚠️ HWND reattach skipped (Outlook window not in EXE session tree)');
            return false;
        }
        return this.bindSessionToOutlookWindowElement(winId);
    }

    /** Replace current session with Desktop Root (sees all top-level HWNDs). */
    private async replaceSessionWithRoot(): Promise<void> {
        try {
            await fetch(`${this.baseUrl}/session/${this.sessionId}`, { method: 'DELETE' });
        } catch {
            /* ignore */
        }
        const r = await undiciFetch(`${this.baseUrl}/session`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                desiredCapabilities: {
                    platformName: 'Windows',
                    deviceName: 'WindowsPC',
                    app: 'Root'
                }
            }),
            dispatcher: OutlookWinAppDriverSession.sessionDispatcher
        });
        const d = (await r.json()) as { sessionId?: string; value?: { message?: string } };
        if (!r.ok || !d.sessionId) {
            throw new Error(`Root session failed: ${d.value?.message ?? r.statusText}`);
        }
        this.sessionId = d.sessionId;
        this.sessionUsesDesktopRoot = true;
        await fetch(`${this.baseUrl}/session/${this.sessionId}/timeouts`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ implicit: 100 })
        });
    }

    /**
     * Desktop Root only — do NOT bind to main Outlook HWND here.
     * Binding main then deleting that session after compose opens was destabilizing the new-mail window.
     */
    private async recoverSessionViaRoot(): Promise<void> {
        console.log(' Trying WinAppDriver Root session to find Outlook on the desktop…');
        await this.replaceSessionWithRoot();
        await new Promise((r) => setTimeout(r, 1500));

        const winId = await this.findOutlookWindowElementId(50, 1500);
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
    private async reattachSessionToComposeWindow(): Promise<void> {
        console.log(' Scoping automation to the compose (message) window for GIPHY…');
        if (!this.sessionUsesDesktopRoot) {
            console.log(
                ' Switching to Desktop Root to locate the compose window (avoids ending a main-Outlook HWND session after new mail is open).'
            );
            await this.replaceSessionWithRoot();
        } else {
            console.log(' Already on Desktop Root — skipping extra session reset before compose.');
        }
        await new Promise((r) => setTimeout(r, 600));

        const winId = await this.findComposeWindowElementId(50, 1000);
        if (!winId) {
            throw new Error(
                'Compose window not found (e.g. Untitled - Message (HTML)). Is it open and visible?'
            );
        }
        const ok = await this.bindSessionToOutlookWindowElement(winId);
        if (!ok) {
            throw new Error('Could not bind WinAppDriver to the compose window HWND');
        }
        console.log(' Session is scoped to the compose (Message) window.');
        await new Promise((r) => setTimeout(r, 1800));
    }

    /** After a ribbon click, the add-in search box is a strong signal. Avoid //*[Name*=GIPHY] — it matches the ribbon button that was already there. */
    private async waitForGiphyPaneHint(timeoutMs = 10_000): Promise<boolean> {
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
                const id = await this.tryFindXpath(xp);
                if (id) {
                    return true;
                }
            }
            await new Promise((r) => setTimeout(r, 450));
        }
        return false;
    }

    private async clickGiphyButtonWithRetry(timeoutMs = 45_000): Promise<void> {
        await this.tryFocusComposeRibbon();
        const deadline = Date.now() + timeoutMs;
        let attempt = 0;

        /** Ribbon add-in buttons: icon is usually upper half of the UIA rect — avoid y=0.5 when bbox includes label. */
        const RIBBON_ICON_Y = 0.42;

        const tryGiphyElement = async (eid: string, label: string): Promise<boolean> => {
            console.log(`🔎 Targeting GIPHY control: ${label}`);
            let clicked = false;
            if (await this.pointerClickCenterOfElement(eid, { yFraction: RIBBON_ICON_Y })) {
                clicked = true;
            } else if (await this.pointerClickCenterOfElement(eid, { yFraction: 0.5 })) {
                clicked = true;
            } else {
                const clickRes = await fetch(`${this.baseUrl}/session/${this.sessionId}/element/${eid}/click`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({})
                });
                clicked = clickRes.ok;
            }
            if (!clicked) {
                return false;
            }
            const paneOk = await this.waitForGiphyPaneHint(10_000);
            if (!paneOk) {
                console.log(
                    '⚠️ Click sent but GIPHY/search UI never showed in UIA — likely missed the icon; retrying…'
                );
            } else {
                console.log(' GIPHY task pane / search UI visible in UIA.');
            }
            return paneOk;
        };

        while (Date.now() < deadline) {
            attempt += 1;
            if (attempt > 1 && attempt % 4 === 0) {
                await this.tryFocusComposeRibbon();
            }

            // Do NOT use findElementByStrategy('name','GIPHY') — it returns Group "GIPHY" before Button "GIPHY".

            for (const { xpath, label } of OutlookWinAppDriverSession.GIPHY_BUTTON_XPATHS) {
                const ids = await this.findAllElementIdsByXpath(xpath);
                if (ids.length > 0) {
                    for (let j = ids.length - 1; j >= 0; j--) {
                        const id = ids[j];
                        if (id && (await tryGiphyElement(id, `${label} [match ${String(j)}]`))) {
                            return;
                        }
                    }
                } else {
                    const found = await this.tryFindXpath(xpath);
                    if (found && (await tryGiphyElement(found, label))) {
                        return;
                    }
                }
            }

            const allNamedButtons = await this.findAllElementIdsByXpath(`//Button[@Name='GIPHY']`);
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

            for (const { xpath, label } of OutlookWinAppDriverSession.GIPHY_LOOSE_XPATHS) {
                const found = await this.tryFindXpath(xpath);
                if (!found) {
                    continue;
                }
                if (await tryGiphyElement(found, label)) {
                    return;
                }
            }

            const a11y = await this.findElementByStrategy('accessibility id', 'GIPHY');
            if (a11y && (await tryGiphyElement(a11y, 'accessibility id'))) {
                return;
            }

            await new Promise((r) => setTimeout(r, 1200));
        }
        throw new Error(
            'GIPHY ribbon click did not open the add-in (search UI never appeared in UIA) within timeout'
        );
    }

    private async clickNewEmailButtonWithRetry(timeoutMs = 45_000): Promise<void> {
        const deadline = Date.now() + timeoutMs;
        while (Date.now() < deadline) {
            for (const { xpath, label } of OutlookWinAppDriverSession.NEW_EMAIL_BUTTON_XPATHS) {
                const eid = await this.tryFindXpath(xpath);
                if (!eid) {
                    continue;
                }
                console.log(`🔎 Found New Email: ${label}`);
                const clickRes = await fetch(`${this.baseUrl}/session/${this.sessionId}/element/${eid}/click`, {
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
        await this.sendCtrlN();
        await new Promise((r) => setTimeout(r, 2000));
    }
}
