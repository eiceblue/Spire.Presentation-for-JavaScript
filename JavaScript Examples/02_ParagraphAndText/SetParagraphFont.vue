<template>
  <span>Click the following button to set paragraph font.</span>
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
        let inputFileName = "Template_Az2.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Load PPT document from the specified input file
        ppt.LoadFromFile(inputFileName);

        // Get the first slide
        let slide = ppt.Slides.get_Item(0);

        // Access the first and second placeholder in the slide and typecasting it as AutoShape
        let tf1 = slide.Shapes.get_Item(0).TextFrame;
        let tf2 = slide.Shapes.get_Item(1).TextFrame;

        // Access the first Paragraph
        let para1 = tf1.Paragraphs.get_Item(0);
        let para2 = tf2.Paragraphs.get_Item(0);

        // Justify the paragraph
        para2.Alignment = wasmModule.TextAlignmentType.Justify;

        // Access the first text range
        let textRange1 = para1.FirstTextRange;
        let textRange2 = para2.FirstTextRange;

        // Define new fonts
        let fd1 = wasmModule.TextFont.Create("Elephant");
        let fd2 = wasmModule.TextFont.Create("Castellar");

        // Assign new fonts to text range
        textRange1.LatinFont = fd1;
        textRange2.LatinFont = fd2;

        // Set font to Bold
        textRange1.Format.IsBold = wasmModule.TriState.True;
        textRange2.Format.IsBold = wasmModule.TriState.False;

        // Set font to Italic
        textRange1.Format.IsItalic = wasmModule.TriState.False;
        textRange2.Format.IsItalic = wasmModule.TriState.True;

        // Set font color
        textRange1.Fill.FillType = wasmModule.FillFormatType.Solid;
        textRange1.Fill.SolidColor.Color = wasmModule.Color.get_Purple();
        textRange2.Fill.FillType = wasmModule.FillFormatType.Solid;
        textRange2.Fill.SolidColor.Color = wasmModule.Color.get_Peru();

        const outputFileName = "SetParagraphFont.pptx";

        // Save to file
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
