html2canvas-asp-vbscript-proxy 0.0.3
=====================

#### Vbscript (asp classic) Proxy html2canvas ####

This script allows you to use html2canvas.js with different servers, ports and protocols (http, https),
preventing to occur "tainted" when exporting the `<canvas>` for image.

###Problem and Solution###
When adding an image that belongs to another domain in `<canvas>` and after that try to export the canvas
for a new image, a security error occurs (actually occurs is a security lock), which can return the error:

> SecurityError: DOM Exception 18
>
> Error: An attempt was made to break through the security policy of the user agent.

### Follow ###

I ask you to follow me or ["star"](https://github.com/brcontainer/html2canvas-asp-vbscript-proxy/star) my repository to track updates

### Usage ###

```html
<!DOCTYPE html>
<html>
	<head>
		<meta charset="utf-8">
		<title>html2canvas asp (vbscript) proxy</title>
		<script src="html2canvas.js"></script>
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
	</head>
	<body>
		<p>
			<img alt="google maps static" src="http://maps.googleapis.com/maps/api/staticmap?center=40.714728,-73.998672&amp;zoom=12&amp;size=400x400&amp;maptype=roadmap&amp;sensor=false">
		</p>
		<p>
			<img alt="facebook https image redirect" src="https://graph.facebook.com/1415773021975267/picture">
		</p>
	</body>
</html>
```

### Others scripting language ###

You do not use ASP Classic, but need html2canvas working with proxy, see other proxies:

* [html2canvas proxy in php](https://github.com/brcontainer/html2canvas-php-proxy)
* [html2canvas proxy in asp.net (csharp)](https://github.com/brcontainer/html2canvas-csharp-proxy)

