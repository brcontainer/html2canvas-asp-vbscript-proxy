html2canvas-asp-vbscript-proxy 0.0.1
=====================

#### Vbscript (asp classic) Proxy html2canvas ####

This script allows you to use html2canvas.js with different servers, ports and protocols (http, https), preventing to occur "tainted" when exporting the "canvas" for image.

### Usage ###

```html
<img alt="google maps static" src="http://maps.googleapis.com/maps/api/staticmap?center=40.714728,-73.998672&zoom=12&size=400x400&maptype=roadmap&sensor=false">
<script>
window.onload = function(){
  html2canvas( [ document.body ], {
  	"proxy":"html2canvasproxy.asp",
		"onrendered": function(canvas) {
			var uridata = canvas.toDataURL("image/png");
			window.open(uridata);
		}
	});
};
</script>
```
