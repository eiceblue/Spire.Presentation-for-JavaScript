<template>
  <span>The example shows how to set the outline and effect for shape in a PPT document. </span>
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

        //Set background Image
        let rect = wasmModule.RectangleF.FromLTRB(0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height);
        ppt.Slides.get_Item(0).Shapes.AppendEmbedImage({shapeType:wasmModule.ShapeType.Rectangle,fileName: ImageFileName,rectangle: rect});
        slide.Shapes.get_Item(0).Line.FillFormat.SolidFillColor.Color = wasmModule.Color.get_FloralWhite();

        //Draw a Rectangle shape
        let shape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle, rectangle:wasmModule.RectangleF.FromLTRB(150, 180, 250, 230)});
        shape.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.Fill.SolidColor.Color = wasmModule.Color.get_SkyBlue();
        //Set outline color
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_Red();
        //Set shadow effect
        let shadow = wasmModule.PresetShadow.Create();
        shadow.ColorFormat.Color = wasmModule.Color.get_LightSkyBlue();
        shadow.Preset = wasmModule.PresetShadowValue.FrontRightPerspective;
        shadow.Distance = 10.0;
        shadow.Direction = 225.0;
        shape.EffectDag.PresetShadowEffect = shadow;

        //Draw a Ellipse shape
        shape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Ellipse,rectangle: wasmModule.RectangleF.FromLTRB(400, 150, 500, 250)});
        shape.Fill.FillType = wasmModule.FillFormatType.Solid;
        shape.Fill.SolidColor.Color = wasmModule.Color.get_SkyBlue();
        //Set outline color
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_Yellow();
        //Set shadow effect
        let glow = wasmModule.GlowEffect.Create();
        glow.ColorFormat.Color = wasmModule.Color.get_LightPink();
        glow.Radius = 20.0;
        shape.EffectDag.GlowEffect = glow;

        // Define the output file name
        const outputFileName = "SetOutlineAndEffectForShape_out.pptx";

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
