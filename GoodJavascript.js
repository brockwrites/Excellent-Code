//This is a good example of a JavaScript callback.  It comes from the K15t Scroll Viewport website.

<script>
var clickLink = function() {
            if (document.getElementById('sv-reader-view-toolbar-disable-readonlyview')) {
                document.getElementById('sv-reader-view-toolbar-disable-readonlyview').click();
            } else {
                window.setTimeout(clickLink, 100);
            }
        };
        clickLink();
</script>
