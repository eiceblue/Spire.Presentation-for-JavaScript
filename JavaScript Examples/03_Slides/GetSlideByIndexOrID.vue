<template>
  <span>Click the following button to access slide by index and shape ID.</span>
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
        let inputFileName = "BlankSample_N.pptx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Load PPT document from the specified input file
        ppt.LoadFromFile(inputFileName);

        //Get slide by index 0
        let slide1 = ppt.Slides.get_Item(0);

        // Append a shape in the slide
        let shape1 = slide1.Shapes.AppendShape({ shapeType: wasmModule.ShapeType.Rectangle, rectangle: wasmModule.RectangleF.FromLTRB(100, 100, 300, 200) });
        
        // Add text in the shape
        shape1.TextFrame.Text = "Get slide by index";

        // Get slide by slide ID
        let slide2 = ppt.FindSlide(Number(ppt.Slides.get_Item(1).SlideID));
        
        // Append a shape in the slide
        let shape2 = slide2.Shapes.AppendShape({ shapeType: wasmModule.ShapeType.Rectangle, rectangle: wasmModule.RectangleF.FromLTRB(100, 100, 300, 200) });
        
        // Add text in the shape
        shape2.TextFrame.Text = "Get slide by slide id";

        const outputFileName = "GetSlideByIndexOrID.pptx";

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
