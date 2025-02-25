<template>
  <span>The example demonstrates how to set shadow effect for the shape in a PPT document.</span>
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

        let slide = ppt.Slides.get_Item(0);

        //Set background image
        let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle,fileName: ImageFileName,rectangle: rect});
        slide.Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

        //Add a shape to slide.
        let rect1 = wasmModule.RectangleF.FromLTRB(200, 150, 500, 270);
        let shape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: rect1});
        shape.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.Fill.SolidColor.Color = wasmModule.Color.get_LightBlue();
        shape.Line.FillType = wasmModule.FillFormatType.None;
        shape.TextFrame.Text = "This demo shows how to apply shadow effect to shape.";
        shape.TextFrame.TextRange.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.TextFrame.TextRange.Fill.SolidColor.Color = wasmModule.Color.get_Black();

        //Create an inner shadow effect through InnerShadowEffect object.
        let innerShadow = wasmModule.InnerShadowEffect.Create();
        innerShadow.BlurRadius = 20;
        innerShadow.Direction = 0;
        innerShadow.Distance = 0;
        innerShadow.ColorFormat.Color = wasmModule.Color.get_Black();

        //Apply the shadow effect to shape.
        shape.EffectDag.InnerShadowEffect = innerShadow;

        // Define the output file name
        const outputFileName = "SetShadowEffectForShape_out.pptx";

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
