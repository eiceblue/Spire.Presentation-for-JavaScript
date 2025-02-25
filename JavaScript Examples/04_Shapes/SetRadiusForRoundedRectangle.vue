<template>
  <span>The example shows how to set radius for the rounded Rectangle.</span>
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

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Get the first slide
        let slide = ppt.Slides.get_Item(0);

        //Insert a rectangle with four round corners and set its radius
        let shape1 = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.RoundCornerRectangle,rectangle: wasmModule.RectangleF.FromLTRB(50, 50, 200, 200)});
        shape1.SetRoundRadius(shape1.Width / 3);

        //Insert a rectangle with one round corner and set its radius
        let shape2 = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.OneRoundCornerRectangle,rectangle: wasmModule.RectangleF.FromLTRB(250, 50, 400, 200)});
        shape2.SetRoundRadius(shape2.Width / 3);

        //Insert a rectangle with one round corner and which one round cornet is snipped and set its radius
        let shape3 = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.OneSnipOneRoundCornerRectangle,rectangle: wasmModule.RectangleF.FromLTRB(450, 50, 600, 200)});
        shape3.SetRoundRadius(shape3.Width / 3);

        //Insert a rectangle with two diagonal round corners and set its radius
        let shape4 = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.TwoDiagonalRoundCornerRectangle,rectangle: wasmModule.RectangleF.FromLTRB(50, 250, 200, 400)});
        shape4.SetRoundRadius(shape4.Width / 3);

        //Insert a rectangle with two same side round corners and set its radius
        let shape5 = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.TwoSamesideRoundCornerRectangle, rectangle:wasmModule.RectangleF.FromLTRB(250, 250, 400, 400)});
        shape5.SetRoundRadius(shape5.Width / 3);

        // Define the output file name
        const outputFileName = "SetRadiusForRoundedRectangle_out.pptx";

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
