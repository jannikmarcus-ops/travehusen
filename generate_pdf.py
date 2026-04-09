from playwright.sync_api import sync_playwright
import os

html_path = os.path.abspath('BPD_Pitch_Travehusen.html')

with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page(viewport={'width': 1920, 'height': 1080})
    page.goto(f'file://{html_path}')
    page.wait_for_load_state('networkidle')

    page.evaluate("""() => {
        document.body.style.overflow = 'visible';
        document.body.style.width = 'auto';
        document.body.style.height = 'auto';
        document.body.style.background = 'white';
        document.documentElement.style.overflow = 'visible';
        document.documentElement.style.width = 'auto';
        document.documentElement.style.height = 'auto';

        const slides = document.querySelectorAll('.slide');
        slides.forEach(s => {
            s.style.position = 'relative';
            s.style.opacity = '1';
            s.style.pointerEvents = 'all';
            s.style.pageBreakAfter = 'always';
            s.style.width = '1920px';
            s.style.height = '1080px';
        });

        const inners = document.querySelectorAll('.slide-inner');
        inners.forEach(si => {
            si.style.transform = 'none';
        });

        const anims = document.querySelectorAll('.anim');
        anims.forEach(a => {
            a.style.opacity = '1';
            a.style.animation = 'none';
        });
    }""")

    page.pdf(
        path='BPD_Pitch_Travehusen.pdf',
        width='1920px',
        height='1080px',
        print_background=True,
        margin={'top': '0', 'right': '0', 'bottom': '0', 'left': '0'}
    )
    browser.close()
    print('PDF erstellt')
