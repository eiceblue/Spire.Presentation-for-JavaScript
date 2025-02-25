<template>
  <span>Click the following button to create mutiple level bullets.</span>
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

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "Bullets2.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Load PPT document from the specified input file
        ppt.LoadFromFile(inputFileName);

        // Get the first slide
        let slide = ppt.Slides.get_Item(0);

        // Retrieve the TextFrame of the second shape
        let tf1 = slide.Shapes.get_Item(1).TextFrame;

        // Access the first Paragraph and set bullet style
        let para = tf1.Paragraphs.get_Item(0);
        para.BulletType = wasmModule.TextBulletType.Symbol;
        para.BulletChar = 8226;
        para.Depth = 0;

        // Access the second Paragraph and set bullet style
        para = tf1.Paragraphs.get_Item(1);
        para.BulletType = wasmModule.TextBulletType.Symbol;
        para.BulletChar = 45;
        para.Depth = 1;

        // Access the third Paragraph and set bullet style
        para = tf1.Paragraphs.get_Item(2);
        para.BulletType = wasmModule.TextBulletType.Symbol;
        para.BulletChar = 8226;
        para.Depth = 2;

        // Access the fourth Paragraph and set bullet style
        para = tf1.Paragraphs.get_Item(3);
        para.BulletType = wasmModule.TextBulletType.Symbol;
        para.BulletChar = 45;
        para.Depth = 3;

        const outputFileName = "MultipleLevelBullets.pptx";

        // Save to file
        ppt.SaveToFile({file:outputFileName,fileFormat:wasmModule.FileFormat.Pptx2013});

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
