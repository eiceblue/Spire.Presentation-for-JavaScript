<template>
  <span>The example demonstrates how to group shapes in a PPT document.</span>
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

        //Get the first slide
        let slide = ppt.Slides.get_Item(0);

        //Set background image
        let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        slide.Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle,fileName: ImageFileName,rectangle: rect});
        slide.Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

        //Create two shapes in the slide
        let rectangle = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: wasmModule.RectangleF.FromLTRB(250, 180, 450, 220)});
        rectangle.Fill.FillType = wasmModule.FillFormatType.Solid;
        rectangle.Fill.SolidColor.KnownColor = wasmModule.KnownColors.SkyBlue;
        rectangle.Line.Width = 0.1;
        let ribbon = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Ribbon2, rectangle:wasmModule.RectangleF.FromLTRB(290, 155, 410, 235)});
        ribbon.Fill.FillType = wasmModule.FillFormatType.Solid;
        ribbon.Fill.SolidColor.KnownColor = wasmModule.KnownColors.LightPink;
        ribbon.Line.Width = 0.1;

        //Add the two shape objects to an array list
        let list = [];
        list.push(rectangle);
        list.push(ribbon);

        //Group the shapes in the list
        ppt.Slides.get_Item(0).GroupShapes(list);

        // Define the output file name
        const outputFileName = "GroupShapes_out.pptx";

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
