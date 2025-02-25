<template>
  <span>This sample demonstrates how to insert shapes into a PPT document.</span>
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

        //Append new shape - Triangle and set style
        let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Triangle, rectangle:wasmModule.RectangleF.FromLTRB(115, 130, 215, 230)});
        shape.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.Fill.SolidColor.Color = wasmModule.Color.get_LightGreen();
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();

        //Append new shape - Ellipse
        shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Ellipse,rectangle: wasmModule.RectangleF.FromLTRB(290, 130, 440, 230)});
        shape.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.Fill.SolidColor.Color = wasmModule.Color.get_LightSkyBlue();
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();

        //Append new shape - Heart
        shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Heart, rectangle:wasmModule.RectangleF.FromLTRB(470, 130, 600, 230)});
        shape.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.Fill.SolidColor.Color = wasmModule.Color.get_Red();
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_LightGray();


        //Append new shape - FivePointedStar
        shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.FivePointedStar, rectangle:wasmModule.RectangleF.FromLTRB(90, 270, 240, 420)});
        shape.Fill.FillType = wasmModule.FillFormatType.Gradient;
        shape.Fill.SolidColor.Color = wasmModule.Color.get_Black();
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();

        //Append new shape - Rectangle
        shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(320, 290, 420, 410)});
        shape.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.Fill.SolidColor.Color = wasmModule.Color.get_Pink();
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_LightGray();

        //Append new shape - BentUpArrow
        shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.BentUpArrow, rectangle:wasmModule.RectangleF.FromLTRB(470, 300, 620, 400)});

        //Set the color of shape
        shape.Fill.FillType = wasmModule.FillFormatType.Gradient;
        shape.Fill.Gradient.GradientStops.Append({position:1,knownColor: wasmModule.KnownColors.Olive});
        shape.Fill.Gradient.GradientStops.Append({position:0,knownColor: wasmModule.KnownColors.PowderBlue});
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();

        // Define the output file name
        const outputFileName = "AddShapes_out.pptx";

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
