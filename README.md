# Spire.Presentation for JavaScript

A powerful presentation processing library for developers to create, read, edit, and convert PowerPoint files in JavaScript.

[![Foo](https://i.imgur.com/r8c1F0l.png)](https://www.e-iceblue.com/Introduce/presentation-for-javascript.html)

[Product Page](https://www.e-iceblue.com/Introduce/presentation-for-javascript.html) | Documentation | Examples | [Forum](https://www.e-iceblue.com/forum/spire-presentation-f14.html) | [Temporary License](https://www.e-iceblue.com/TemLicense.html) | [Customized Demo](https://www.e-iceblue.com/Misc/customized-demo.html)

[**Spire.Presentation for JavaScript**](https://www.e-iceblue.com/Introduce/presentation-for-javascript.html) is a comprehensive library that empowers developers to seamlessly integrate presentation functionalities into their web applications, providing a robust JavaScript solution for generating, editing, and managing PowerPoint files without the need for Microsoft Office.

This library offers an extensive feature set and supports a variety of PowerPoint formats, including PPT, PPS, PPTX, PPSX, and more. Using it, you can easily handle text, images, shapes, tables, animations, audio, and video on your slides. Additionally, it allows high-quality conversions from PowerPoint presentations to PDF, HTML, XPS, and images (PNG, JPG, TIFF, EMF, SVG).

Spire.Presentation for JavaScript is fully compatible with popular JavaScript frameworks such as Vue, React and Angular, enabling developers to build feature-rich applications that run smoothly across different platforms and browsers, which is an ideal choice for streamlining presentation creation and managing workflows.


### 100% Standalone JavaScript API - No Microsoft Office Needed
Spire.Presentation for JavaScript is a completely independent JavaScript library for processing PowerPoint presentations, eliminating the need for Microsoft Office to be installed on systems. It delivers a fast and reliable solution for managing large presentations, making it a more efficient alternative to traditional Microsoft Office Automation methods.

### High-Quality File Conversion 
Spire.Presentation for JavaScript allows high-quality conversion from PowerPoint files to popular formats such as converting PPT/PPTX to PDF, HTML, XPS, images (PNG, JPG, BMP, SVG), and also supports interconversion between PowerPoint presentation formats.

### Extensive Presentation Processing Capabilities
Spire.Presentation for JavaScript supports various elements in PowerPoint files. With it, you can create slides, add text, images, shapes, tables, charts, watermarks, headers and footers, annotations, notes, SmartArt, audio and video. In addition, interactive elements such as hyperlinks and animations can also be added to your presentations to make them more engaging and dynamic.

### Rich Support for Presentation Formats
Spire.Presentation for JavaScript supports Microsoft PowerPoint 97-2003 and Microsoft PowerPoint 2007, 2010.
● PPT PowerPoint Presentation 97-2003
● PPS PowerPoint SlideShow 97-2003
● PPTX PowerPoint Presentation 2007, 2010, 2013, 2016 and 2019
● PPSX PowerPoint SlideShow 2007, 2010

### Cross-Platform Compatibility
Compatibility is a critical factor when it comes to web applications. Spire.Presentation for JavaScript provides robust support for major frameworks and web-based environments, and applications built with it can run seamlessly across different browsers and platforms, providing a consistent user experience.

### Improve Developer Productivity
Spire.Presentation for JavaScript is easy to use and requires no additional tools or software. Developers using it can easily incorporate presentation functionalities into their projects to automate document operations, reducing development time and effort.

## Vue Examples

### Convert PowerPoint to PDF in JavaScript
```JavaScript
<template>
  <span>The following example demonstrates how to convert a PowerPoint document to PDF</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);
        
        // Load the input file into the virtual file system (VFS)
        const inputFileName = "ToPDF.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt =wasmModule.Presentation.Create();

        //Load PPT file from disk
        ppt.LoadFromFile(inputFileName);

        // Define the output file name
        const outputFileName = "ToPDF.pdf";

        // Save the document 
        ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.PDF });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/pdf" });

        // Clean up resources
        ppt.Dispose();

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
```

### Convert PowerPoint to HTML in JavaScript
```JavaScript
<template>
  <span>The following example demonstrates how to convert a PowerPoint document to HTML</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);
         
        // Load the input file into the virtual file system (VFS)
        const inputFileName = "Conversion.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create an instance of presentation document
        let ppt =wasmModule.Presentation.Create();

        // Load file
        ppt.LoadFromFile(inputFileName);

        // Define the output file name
        const outputFileName = "ToHTML.html";

        // Save the document 
        ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/html" });

        // Clean up resources
        ppt.Dispose();

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
```

### Add images in PowerPoint slides in JavaScript
```JavaScript
<template>
  <span>The sample demonstrates how to insert image into a PowerPoint document</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);
 
        // Load the input file and image into the virtual file system (VFS)
        const inputFileName = "InsertImage.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);
        const ImageFileName = "InsertImage.png";
        await wasmModule.FetchFileToVFS(ImageFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create PPT document and load file
        let ppt = wasmModule.Presentation.Create();
        ppt.LoadFromFile(inputFileName);

        //Insert image to PPT
        let rect1 = wasmModule.RectangleF.FromLTRB(ppt.SlideSize.Size.Width / 2 - 280, 140, (120 + ppt.SlideSize.Size.Width / 2 - 280), 260);
        let image = ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({ shapeType: wasmModule.ShapeType.Rectangle, fileName: ImageFileName, rectangle: rect1 });
        image.Line.FillType = wasmModule.FillFormatType.None;

        // Define the output file name
        const outputFileName = "InsertImage.pptx";

        // Save the document 
        ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013 });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });

        // Clean up resources
        ppt.Dispose();

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
```

[Product Page](https://www.e-iceblue.com/Introduce/presentation-for-javascript.html) | Documentation | Examples | [Forum](https://www.e-iceblue.com/forum/spire-presentation-f14.html) | [Temporary License](https://www.e-iceblue.com/TemLicense.html) | [Customized Demo](https://www.e-iceblue.com/Misc/customized-demo.html)

