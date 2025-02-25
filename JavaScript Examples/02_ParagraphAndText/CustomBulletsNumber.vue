<template>
  <span>Click the following button to customize bullets number.</span>
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

        // Access the first placeholder in the slide and typecasting it as AutoShape
        let tf1 = slide.Shapes.get_Item(1).TextFrame;

        // Access the first Paragraph and set bullet style
        let para = tf1.Paragraphs.get_Item(0);
        para.Depth = 0;
        para.BulletType = wasmModule.TextBulletType.Numbered;
        para.BulletStyle = wasmModule.NumberedBulletStyle.BulletArabicPeriod;
        para.BulletNumber = 2;

        // Access the second Paragraph and set bullet style
        para = tf1.Paragraphs.get_Item(1);
        para.Depth = 0;
        para.BulletType = wasmModule.TextBulletType.Numbered;
        para.BulletStyle = wasmModule.NumberedBulletStyle.BulletArabicPeriod;
        para.BulletNumber = 4;

        // Access the third Paragraph and set bullet style
        para = tf1.Paragraphs.get_Item(2);
        para.Depth = 0;
        para.BulletType = wasmModule.TextBulletType.Numbered;
        para.BulletStyle = wasmModule.NumberedBulletStyle.BulletArabicPeriod;
        para.BulletNumber = 6;

        // Access the fourth Paragraph and set bullet style
        para = tf1.Paragraphs.get_Item(3);
        para.Depth = 0;
        para.BulletType = wasmModule.TextBulletType.Numbered;
        para.BulletStyle = wasmModule.NumberedBulletStyle.BulletArabicPeriod;
        para.BulletNumber = 7;

        const outputFileName = "CustomBulletsNumber.pptx";

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
