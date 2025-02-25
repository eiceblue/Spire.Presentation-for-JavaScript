<template>
  <span>The example demonstrates how to reset shape size and position in a PPT document.</span>
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

        const inputFileName = "ShapeTemplate.pptx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Load file
        ppt.LoadFromFile(inputFileName);

        //Define the original slide size
        let currentHeight = ppt.SlideSize.Size.Height;
        let currentWidth = ppt.SlideSize.Size.Width;

        //Change the slide size as A3
        ppt.SlideSize.Type = wasmModule.SlideSizeType.A3;

        //Define the new slide size
        let newHeight = ppt.SlideSize.Size.Height;
        let newWidth = ppt.SlideSize.Size.Width;

        //Define the ratio from the old and new slide size
        let ratioHeight = newHeight / currentHeight;
        let ratioWidth = newWidth / currentWidth;

        //Reset the size and position of the shape on the slide
        for (let i = 0; i < ppt.Slides.Count; i++) {
            let slide = ppt.Slides.get_Item(i);
            for (let j = 0; j < slide.Shapes.Count; j++) {
                let shape = slide.Shapes.get_Item(j);
                shape.Height = shape.Height * ratioHeight;
                shape.Width = shape.Width * ratioWidth;

                shape.Left = shape.Left * ratioHeight;
                shape.Top = shape.Top * ratioWidth;
            }
        }

        // Define the output file name
        const outputFileName = "ResetShapeSizeAndPosition_out.pptx";

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
