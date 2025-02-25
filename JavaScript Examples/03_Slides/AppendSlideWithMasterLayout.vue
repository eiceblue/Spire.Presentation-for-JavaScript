<template>
  <span>Click the following button to append a new slide with master layout in a PPT document.</span>
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
        let inputFileName = "AppendSlideWithMasterLayout.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Load PPT document from the specified input file
        ppt.LoadFromFile(inputFileName);

        // Get the master
        let master = ppt.Masters.get_Item(0);

        // Get master layout slides
        let masterLayouts = master.Layouts;
        let layoutSlide = masterLayouts.get_Item(1);

        // Append a rectangle to the layout slide
        let shape = layoutSlide.Shapes.AppendShape({ shapeType: wasmModule.ShapeType.Rectangle, rectangle: spirepresentation.RectangleF.FromLTRB(10, 50, 110, 130) });

        // Add a text into the shape and set the style
        shape.Fill.FillType = wasmModule.FillFormatType.None;
        shape.AppendTextFrame("Layout slide 1");
        shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Arial Black");
        shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.get_CadetBlue();

        // Append new slide with master layout
        ppt.Slides.Append({ slide: ppt.Slides.get_Item(0), layout: master.Layouts.get_Item(1) });

        // Another way to append new slide with master layout
        ppt.Slides.Insert({ index: 2, slide: ppt.Slides.get_Item(1), layout: master.Layouts.get_Item(1) });

        const outputFileName = "AppendSlideWithMasterLayout.pptx";

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
