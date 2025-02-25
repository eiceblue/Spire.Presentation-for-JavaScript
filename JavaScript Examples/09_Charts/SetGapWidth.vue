<template>
    <span>The following example demonstrates how to set gap width for chart in a PPT document</span>
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
                
                // Load the input file into the virtual file system (VFS)
                const inputFileName = "ChartSample2.pptx";
                await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

                // Create PPT document and load file
                let ppt = wasmModule.Presentation.Create();
                
                // Load file from VFS
                ppt.LoadFromFile(inputFileName);

                // Get chart on the first slide
                let Chart = ppt.Slides.get_Item(0).Shapes.get_Item(0);

                // Set gap width
                Chart.GapWidth = 50;

                // Save to file
                const outputFileName = "SetGapWidth.pptx";
                ppt.SaveToFile({file:outputFileName,fileFormat:spirepresentation.FileFormat.Pptx2010});

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