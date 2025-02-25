<template>
  <div class="app-description">
    <h1 class="title">Welcome to the Vue Demo</h1>
    <p>
      We are excited to introduce our Vue application, showcasing the powerful capabilities of Spire.Presentation for JavaScript in handling presentations files. This project seamlessly integrates the functionalities of Spire.Presentation for JavaScript through modern web application frameworks. 
    </p>
    <p class="title">Getting Started</p>
    <ol>
      <li>Copy the two files (<span class="getstart">Spire.Presentation.Base.js & Spire.Presentation.Base.wasm</span>) from <span class="getstart">lib</span> directory to the <span class="getstart">public</span> directory in Vue project.</li>
      <li>Please download the <b>ARIALUNI.TTF</b> font and upload it to the <span class="getstart">public/static/font</span> folder in your project directory. </li>
        <p><b>Note:</b>The ARIALUNI.TTF font is protected by copyright and is only to be used legally for personal and commercial purposes.</p>If your document uses additional fonts, you can also upload these fonts to the public/static/font directory. After uploading, ensure that the font names in your code are adjusted accordingly.
      <li>Expand Spire.Presentation from the left menu, select a detailed feature and click the "Start" button to execute.</li>
    </ol>
    <p class="title">Product Introduction</p>
    <p>	  
	  Spire.Presentation for JavaScript is a powerful presentation processing library designed for developers in JavaScript applications. It offers extensive functionality for managing presentations, such as create, edit, and customize PowerPoint presentations in JavaScript. You can easily handle text, images, shapes, tables, animations, audio, and video on your slides. Additionally, it supports exporting presentation slides to multiple formats, including images (PNG, JPG, TIFF, EMF, SVG), PDF, XPS, and HTML. This versatility makes it an ideal choice for developers looking to create dynamic and engaging presentations.	  
    </p>
    <p class="title">Technical Architecture</p>
    <p>
      Built using WebAssembly technology, Spire.Presentation for JavaScript is fully compatible with popular frameworks such as Vue, React, Angular, and plain JavaScript, 
      enabling developers to easily create and distribute their JavaScript applications on the web.
    </p>
    <p class="title">File Support</p>
    <p>
      This versatile JavaScript presentations API operates independently, with no reliance on Microsoft Office presentations. 
      Spire.Presentation for JavaScript supports Microsoft PowerPoint 97-2003 and Microsoft PowerPoint 2007, 2010, 2016, and 2019 (PPTX, PPSX). Compared to custom solutions or Microsoft Automation, Spire.Presentation for JavaScript offers superior performance, reliability, and ease of use.
    </p>
    <p class="title">Contact Us</p>
    <ul>
      <li>For any technical questions regarding the product, feel free to contact us: <a href="mailto:support@e-iceblue.com">support@e-iceblue.com</a>.</li>
      <li>For any questions about purchasing the product, please reach out to: <a href="mailto:sales@e-iceblue.com">sales@e-iceblue.com</a>.</li>
      <li>If you can't find the example you need, contact us for a free custom demo: <a href="mailto:support@e-iceblue.com">support@e-iceblue.com</a>.</li>
    </ul> 
  </div>
</template>
<script>
export default {
  name: 'AppDescription',
};
</script>

<style scoped>
.app-description {
  padding: 20px;
  font-family: Arial, sans-serif;
  text-align: left;
}
.title {
  text-align: left;
  font-weight: bold;
  color: #2980B9;
  font-size: 17px;
  padding: 0;
  line-height: 1.4;
}

h1 {
  color: #1F618D;
}

p {
  line-height: 1.4;
}
.getstart
{
  color: #E67E22 ;
}
ul {
  list-style-type: none;
  padding: 0;
}
ol {
  list-style-type: style type num;;
  padding: 0;
}
b
{
  font-weight: bold
};
ul li{
  margin: 10px 0;
}
</style>