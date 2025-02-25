<template>
  <span>Click the following button to add hyperlink to image in PowerPoint document.</span>
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

        let inputFileName = "Template_Ppt_5.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        let imageFileName = "Logo1.png";
        await wasmModule.FetchFileToVFS(imageFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Get the first slide.
        let slide = ppt.Slides.get_Item(0);

        let rect = wasmModule.RectangleF.FromLTRB(480, 350, 640, 510);
        let image = slide.Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle,fileName: imageFileName,rectangle: rect});

        //Add hyperlink to the image.
        let hyperlink = wasmModule.ClickHyperlink.Create("https://www.e-iceblue.com");
        image.Click = hyperlink;

        // Define the output file name
        const outputFileName = "AddHyperlinkToImage_out.pptx";

        // Save the document to the specified path
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
