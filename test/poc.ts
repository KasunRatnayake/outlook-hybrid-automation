import { OutlookCdpGiphyHelper } from './helpers/OutlookCdpGiphyHelper';
import { OutlookWinAppDriverSession } from './helpers/OutlookWinAppDriverSession';

/** Set `POC_SKIP_CDP=1` to finish after GIPHY opens (no Playwright). */
const POC_SKIP_CDP = process.env.POC_SKIP_CDP === '1';

/**
 * `desktop` / `1` / `part1` = WinAppDriver through GIPHY pane only.
 * `webview` / `2` / `part2` = Playwright/CDP only (Outlook + GIPHY must already be open).
 * Omit = full flow.
 */
type PocPart = 'full' | 'desktop' | 'webview';

function resolvePocPart(): PocPart {
    const p = process.env.POC_PART?.trim().toLowerCase();
    if (p === 'webview' || p === '2' || p === 'part2') return 'webview';
    if (p === 'desktop' || p === '1' || p === 'part1') return 'desktop';
    return 'full';
}

/**
 * Hybrid POC: Part 1 = WinAppDriver (Outlook → GIPHY pane). Part 2 = CDP / Playwright (search + GIF click).
 */
async function runHybridPOC(): Promise<void> {
    const pocPart = resolvePocPart();
    const cdp = new OutlookCdpGiphyHelper();

    if (pocPart === 'webview') {
        try {
            await cdp.runGiphyWebViewFlow();
            console.log('🎉 SUCCESS (Part 2 — WebView only).');
        } catch (e) {
            console.error(' Failed:', e instanceof Error ? e.message : String(e));
        }
        return;
    }

    const session = await OutlookWinAppDriverSession.start();
    try {
        await session.runDesktopFlowToGiphyPane();

        if (POC_SKIP_CDP) {
            console.log(
                ' POC_SKIP_CDP=1 — skipping Playwright (WebView2 in Outlook rarely exposes CDP on 9222 without extra setup).'
            );
            console.log('🎉 SUCCESS (WinAppDriver + GIPHY side panel).');
        } else if (pocPart === 'desktop') {
            console.log(
                '🎉 SUCCESS (Part 1 — Outlook + GIPHY pane). Run Part 2: POC_PART=webview npx ts-node test/poc.ts'
            );
        } else {
            await cdp.runGiphyWebViewFlow();
            console.log('🎉 SUCCESS (full hybrid: desktop + Playwright).');
        }
    } catch (e) {
        console.error(' Failed:', e instanceof Error ? e.message : String(e));
    } finally {
        await session.dispose();
    }
}

void runHybridPOC();
