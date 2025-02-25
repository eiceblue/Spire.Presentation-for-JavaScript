<template>
  <span>The example demonstrates how to add round corner rectangle in a PPT document. </span>
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
        
        //Set background image
        const ImageFileName = "bg.png";
        await wasmModule.FetchFileToVFS(ImageFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        let rect =  wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        ppt.Slides.get_Item(0).Shapes.AppendEmbedImage( {shapeType:wasmModule.ShapeType.Rectangle, fileName:ImageFileName, rectangle:rect});
        ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color =  wasmModule.Color.get_FloralWhite();

        //Append a round corner rectangle and set its radius
        let shape = ppt.Slides.get_Item(0).Shapes.AppendRoundRectangle(300, 90, 100, 200, 80);
        //Set the color and fill style of shape
        shape.Fill.FillType =  wasmModule.FillFormatType.Solid;
        shape.Fill.SolidColor.Color =  wasmModule.Color.get_LightBlue();
        shape.ShapeStyle.LineColor.Color =  wasmModule.Color.get_SkyBlue();
        //Rotate the shape to 90 degree
        shape.Rotation = 90;

        // Define the output file name
        const outputFileName = "AddRoundCornerRectangle_out.pptx";

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
