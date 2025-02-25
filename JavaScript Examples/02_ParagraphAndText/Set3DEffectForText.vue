<template>
  <span>Click the following button to set 3D effect for the text in a PPT document.</span>
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

        // Create a PPT document
        let ppt = wasmModule.Presentation.Create();

        // Get the first slide
        let slide = ppt.Slides.get_Item(0);

        // Append a new shape to slide and set the line color and fill type
        let shape = slide.Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: wasmModule.RectangleF.FromLTRB(30, 40, 680, 240)});
        shape.ShapeStyle.LineColor.Color = wasmModule.Color.get_White();
        shape.Fill.FillType = wasmModule.FillFormatType.None;

        // Add text to the shape
        shape.AppendTextFrame("This demo shows how to add 3D effect text to Presentation slide");

        // Set the color of text in shape
        let textRange = shape.TextFrame.TextRange;
        textRange.Fill.FillType = wasmModule.FillFormatType.Solid;
        textRange.Fill.SolidColor.Color = wasmModule.Color.get_LightBlue();

        // Set the Font of text in shape
        textRange.FontHeight = 40;
        textRange.LatinFont = wasmModule.TextFont.Create("Gulim");

        // Set 3D effect for text
        shape.TextFrame.TextThreeD.ShapeThreeD.PresetMaterial = wasmModule.PresetMaterialType.Matte;
        shape.TextFrame.TextThreeD.LightRig.PresetType = wasmModule.PresetLightRigType.Sunrise;
        shape.TextFrame.TextThreeD.ShapeThreeD.TopBevel.PresetType = wasmModule.BevelPresetType.Circle;
        shape.TextFrame.TextThreeD.ShapeThreeD.ContourColor.Color = wasmModule.Color.get_Green();
        shape.TextFrame.TextThreeD.ShapeThreeD.ContourWidth = 3;

        const outputFileName = "Set3DEffectForText.pptx";

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
