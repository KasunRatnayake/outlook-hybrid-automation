import { chromium, type Page } from '@playwright/test';

const OUTLOOK_CDP_URL = process.env.OUTLOOK_CDP_URL ?? 'http://localhost:9222';

/**
 * Playwright over Chrome DevTools (WebView2): connect, search GIPHY, click first result (Part 2).
 */
export class OutlookCdpGiphyHelper {
    private resolveCdpUrlsToTry(): string[] {
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

    private async probeCdpHttpRoot(rootUrl: string): Promise<boolean> {
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

    private logWebView2CdpInstructions(): void {
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

    private async connectCdpWithRetry(
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

    private async connectPlaywrightToOutlookWebViewCdp(): Promise<
        Awaited<ReturnType<typeof chromium.connectOverCDP>>
    > {
        const urls = this.resolveCdpUrlsToTry();
        let anyHttp = false;
        for (const u of urls) {
            if (await this.probeCdpHttpRoot(u)) {
                anyHttp = true;
                console.log(` CDP HTTP probe OK (${u}) — DevTools endpoint is up.`);
                break;
            }
        }
        if (!anyHttp) {
            console.log(
                ' CDP probe: no response on /json/version — WebView2 is almost certainly running without --remote-debugging-port.'
            );
            this.logWebView2CdpInstructions();
        }

        const errors: string[] = [];
        for (const url of urls) {
            try {
                console.log(` Trying Playwright connectOverCDP → ${url}`);
                return await this.connectCdpWithRetry(url, 6, 1200);
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

    private async clickFirstGiphySearchResult(page: Page): Promise<void> {
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

    /** Connect CDP, search "cat", click first GIF (compose + GIPHY pane must be open). */
    async runGiphyWebViewFlow(): Promise<void> {
        console.log(' Part 2 — Playwright (CDP / GIPHY WebView)…');
        console.log('🔗 Connecting Playwright over CDP (WebView2 / Chromium DevTools)…');
        const browser = await this.connectPlaywrightToOutlookWebViewCdp();
        const page = browser.contexts()[0].pages().find((p) => p.url().includes('giphy'));

        if (!page) throw new Error('Giphy WebView not found. Is the add-in pane open?');

        await page.fill('input[placeholder*="Search"]', 'cat');
        await page.keyboard.press('Enter');
        await this.clickFirstGiphySearchResult(page);
    }
}
