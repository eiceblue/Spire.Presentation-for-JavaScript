<template>
  <span>The example demonstrates how to add line with arrow in a PPT document.</span>
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
        const ImageFileName = "bg.png";
        await wasmModule.FetchFileToVFS(ImageFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Set background image
        let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle,fileName: ImageFileName,rectangle: rect});
        ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

        //Add a line to the slides and set its color to red
        let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Line, rectangle:wasmModule.RectangleF.FromLTRB(150, 100, 250, 200)});
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_Red();
        //Set the line end type as StealthArrow
        shape.Line.LineEndType = wasmModule.LineEndType.StealthArrow;

        //Add a line to the slides and use default color
        shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Line, rectangle:wasmModule.RectangleF.FromLTRB(300, 150, 400, 250)});
        shape.Rotation = -45;
        //Set the line end type as TriangleArrowHead
        shape.Line.LineEndType = wasmModule.LineEndType.TriangleArrowHead;

        //Add a line to the slides and set its color to Green
        shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Line,rectangle: wasmModule.RectangleF.FromLTRB(450, 100, 550, 200)});
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_Green();
        shape.Rotation = 90;
        //Set the line begin type as TriangleArrowHead
        shape.Line.LineBeginType = wasmModule.LineEndType.StealthArrow;

        // Define the output file name
        const outputFileName = "AddLineWithArrow_out.pptx";

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
