<template>
  <span>Click the following button to set shadow effect for the text in a PPT document.</span>
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

        // Load the image file into the virtual file system (VFS)
        let ImageFileName = "bg.png";
        await wasmModule.FetchFileToVFS(ImageFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Create a rectangle using the specified coordinates
        let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        
        // Append an embedded image as a rectangle shape to the first slide
        ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({shapeType: wasmModule.ShapeType.Rectangle,fileName: ImageFileName, rectangle:rect});
        
        // Set the line color of the first shape (the embedded image) to a floral white color
        ppt.Slides.get_Item(0).Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

        // Get reference of the slide
        let slide = ppt.Slides.get_Item(0);

        // Add a new rectangle shape to the first slide
        let shape = slide.Shapes.AppendShape({shapeType: wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(120, 100, 570, 300)});
        shape.Fill.FillType = wasmModule.FillFormatType.None;

        // Add the text to the shape and set the font for the text
        shape.AppendTextFrame("Text shading on slides");
        shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).LatinFont = wasmModule.TextFont.Create("Arial Black");
        shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).FontHeight = 21;
        shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.TextFrame.Paragraphs.get_Item(0).TextRanges.get_Item(0).Fill.SolidColor.Color = wasmModule.Color.get_Black();

        ////Add inner shadow and set all necessary parameters
        //InnerShadowEffect Shadow = InnerShadowEffect();

        //Add outer shadow and set all necessary parameters
        let Shadow = wasmModule.OuterShadowEffect.Create();

        Shadow.BlurRadius = 0;
        Shadow.Direction = 50;
        Shadow.Distance = 10;
        Shadow.ColorFormat.Color = wasmModule.Color.get_LightBlue();

        //shape.TextFrame.TextRange.EffectDag.InnerShadowEffect = Shadow;
        shape.TextFrame.TextRange.EffectDag.OuterShadowEffect = Shadow;

        const outputFileName = "SetShadowEffect.pptx";

        // Save to file
        ppt.SaveToFile({file:outputFileName,fileFormat:wasmModule.FileFormat.Pptx2013});

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
