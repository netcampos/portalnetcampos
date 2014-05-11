

<script type="text/javascript">
function downloadJSAtOnload() {var hash = Math.round(Math.random()*1000000); var element = document.createElement("script");element.src = "<%=_lib.domainAssets()%>/js/head.js,jquery.js,bootstrap.js,app.js&t="+hash;document.body.appendChild(element);}
if (window.addEventListener){window.addEventListener("load", downloadJSAtOnload, false);}else if (window.attachEvent){window.attachEvent("onload", downloadJSAtOnload);}else {window.onload = downloadJSAtOnload;}
</script>
</body>
</html>
