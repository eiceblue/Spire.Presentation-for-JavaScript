<template>
  <span>The example demonstrates how to set the format for lines in a PPT document. </span>
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

        const ImageFileName = "bg.png";
        await wasmModule.FetchFileToVFS(ImageFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        //Set background image
        let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle,fileName: ImageFileName,rectangle: rect});
        ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

        //Add a rectangle shape to the slide
        let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: wasmModule.RectangleF.FromLTRB(100, 150, 300, 250)});
        //Set the fill color of the rectangle shape
        shape.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.Fill.SolidColor.Color = wasmModule.Color.get_White();
        //Apply some formatting on the line of the rectangle
        shape.Line.Style = wasmModule.TextLineStyle.ThickThin;
        shape.Line.Width = 5;
        shape.Line.DashStyle = wasmModule.LineDashStyleType.Dash;
        //Set the color of the line of the rectangle
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_SkyBlue();

        //Add a ellipse shape to the slide
        shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Ellipse,rectangle: wasmModule.RectangleF.FromLTRB(400, 150, 600, 250)});
        //Set the fill color of the ellipse shape
        shape.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.Fill.SolidColor.Color = wasmModule.Color.get_White();
        //Apply some formatting on the line of the ellipse
        shape.Line.Style = wasmModule.TextLineStyle.ThickBetweenThin;
        shape.Line.Width = 5;
        shape.Line.DashStyle = wasmModule.LineDashStyleType.DashDot;
        //Set the color of the line of the ellipse
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_OrangeRed();

        // Define the output file name
        const outputFileName = "SetFormatForLines_out.pptx";

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
