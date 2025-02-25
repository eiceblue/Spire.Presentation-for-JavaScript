<template>
  <span>The sample demonstrates how to add custom path animation in PPT.</span>
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

        //Add shape
        let shape = ppt.Slides.get_Item(0).Shapes.AppendShape({shapeType:wasmModule.ShapeType.Rectangle,rectangle: wasmModule.RectangleF.FromLTRB(0, 0, 200, 200)});

        //Add animation
        let effect = ppt.Slides.get_Item(0).Timeline.MainSequence.AddEffect(shape, wasmModule.AnimationEffectType.PathUser);

        let common = effect.CommonBehaviorCollection;

        let motion = common.get_Item(0);
        motion.Origin = wasmModule.AnimationMotionOrigin.Layout;
        motion.PathEditMode = wasmModule.AnimationMotionPathEditMode.Relative;

        //Add moin path
        let moinPath = wasmModule.MotionPath.Create();
        moinPath.Add(wasmModule.MotionCommandPathType.MoveTo, wasmModule.PointF.Create(0,0) , wasmModule.MotionPathPointsType.CurveAuto, true);
        moinPath.Add(wasmModule.MotionCommandPathType.LineTo, wasmModule.PointF.Create(0.1,0.1), wasmModule.MotionPathPointsType.CurveAuto, true);
        moinPath.Add(wasmModule.MotionCommandPathType.LineTo, wasmModule.PointF.Create(-0.1,0.2), wasmModule.MotionPathPointsType.CurveAuto, true);
        moinPath.Add(wasmModule.MotionCommandPathType.End, wasmModule.PointF.Create(0,0), wasmModule.MotionPathPointsType.CurveStraight, true);
        motion.Path = moinPath;

        // Define the output file name
        const outputFileName = "CustomPathAnimation_out.pptx";

        // Save the document to the specified path
        ppt.SaveToFile({ file: outputFileName, fileFormat: wasmModule.FileFormat.Pptx2013});

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
