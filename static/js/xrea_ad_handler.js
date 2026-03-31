(function () {

    let scrolled = false;

    function forceLeftAlign() {
        const bodyChildren = document.body.children;
        for (let i = 0; i < Math.min(bodyChildren.length, 5); i++) {
            const el = bodyChildren[i];
            if (el.classList.contains('page-wrapper')) continue;
            if (el.tagName === 'CENTER' || el.tagName === 'DIV' || el.getAttribute('align') === 'center') {
                el.style.textAlign = 'left';
                el.style.marginLeft = '0';
                el.style.marginRight = 'auto';
                el.style.paddingLeft = '40px';
                el.style.display = 'block';
                const sub = el.querySelectorAll('img, iframe, div');
                sub.forEach(s => {
                    s.style.marginLeft = '0';
                    s.style.marginRight = 'auto';
                    s.style.display = 'inline-block';
                });
                observeAndScroll(el);
            }
        }
    }

    function observeAndScroll(el) {
        if (scrolled) return;

        const ro = new ResizeObserver(entries => {
            for (const entry of entries) {
                const rect = entry.target.getBoundingClientRect();

                // 高さが一定以上＝広告描画完了とみなす
                if (rect.height > 50) {
                    const scrollTarget = window.pageYOffset + (rect.bottom * 2) + 10;

                    window.scrollTo({
                        top: scrollTarget,
                        behavior: 'smooth'
                    });

                    scrolled = true;
                    ro.disconnect();
                }
            }
        });

        ro.observe(el);
    }

    forceLeftAlign();
    const observer = new MutationObserver(forceLeftAlign);
    observer.observe(document.body, { childList: true });

})();
