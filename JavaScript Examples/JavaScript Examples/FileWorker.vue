<template>
  <div class="worker_page">
    <div class="tree_item">
      <SpireTree @item-id="handleChildID" />
    </div>
    <div class="file_item">
      <div class="file-reader">
        <div v-if="filePageId == '0'" class="non-centered">
          <Description />
        </div>
        <!--01_QuickGuide -->
        <div v-if="filePageId == '01-1'">
          <HelloWorld />
        </div>
        <!--02_ParagraphAndText -->
        <div v-if="filePageId == '02-1'">
          <AddParagraph />
        </div>
        <div v-if="filePageId == '02-2'">
          <Alignment />
        </div>
        <div v-if="filePageId == '02-3'">
          <AppendHTML />
        </div>
        <div v-if="filePageId == '02-4'">
          <AutoFitTextOrShape />
        </div>
        <div v-if="filePageId == '02-5'">
          <BordersAndShading />
        </div>
        <div v-if="filePageId == '02-6'">
          <Bullets />
        </div>
        <div v-if="filePageId == '02-7'">
          <ChangeTextStyle />
        </div>
        <div v-if="filePageId == '02-8'">
          <CopyParagraphToAnotherPPT />
        </div>
        <div v-if="filePageId == '02-9'">
          <CustomBulletsNumber />
        </div>
        <div v-if="filePageId == '02-10'">
          <EditPromptText />
        </div>
        <div v-if="filePageId == '02-11'">
          <ExtractText />
        </div>
        <div v-if="filePageId == '02-12'">
          <GetTextFrameEffectiveData />
        </div>
        <div v-if="filePageId == '02-13'">
          <GetTextStyleEffectiveData />
        </div>
        <div v-if="filePageId == '02-14'">
          <HighlightSpecifiedText />
        </div>
        <div v-if="filePageId == '02-15'">
          <Indent />
        </div>
        <div v-if="filePageId == '02-16'">
          <LineSpacing />
        </div>
        <div v-if="filePageId == '02-17'">
          <MixFontStyles />
        </div>
        <div v-if="filePageId == '02-18'">
          <ModifyStyleOfFirstFoundText />
        </div>
        <div v-if="filePageId == '02-19'">
          <MultipleLevelBullets />
        </div>
        <div v-if="filePageId == '02-20'">
          <MultipleParagraphs />
        </div>
        <div v-if="filePageId == '02-21'">
          <PictureCustomBulletStyle />
        </div>
        <div v-if="filePageId == '02-22'">
          <RemoveTextBox />
        </div>
        <div v-if="filePageId == '02-23'">
          <ReplaceText />
        </div>
        <div v-if="filePageId == '02-24'">
          <ReplaceTextRetentionStyle />
        </div>
        <div v-if="filePageId == '02-25'">
          <ReplaceTextWithRegex />
        </div>
        <div v-if="filePageId == '02-26'">
          <RotateText />
        </div>
        <div v-if="filePageId == '02-27'">
          <Set3DEffectForText />
        </div>
        <div v-if="filePageId == '02-28'">
          <SetAnchorOfTextFrame />
        </div>
        <div v-if="filePageId == '02-29'">
          <SetColumnsCountOfTextFrame />
        </div>
        <div v-if="filePageId == '02-30'">
          <SetCustomFonts />
        </div>
        <div v-if="filePageId == '02-31'">
          <SetParagraphFont />
        </div>
        <div v-if="filePageId == '02-32'">
          <SetRightToLeftColumns />
        </div>
        <div v-if="filePageId == '02-33'">
          <SetShadowEffectForText />
        </div>
        <div v-if="filePageId == '02-34'">
          <SetTextDirection />
        </div>
        <div v-if="filePageId == '02-35'">
          <SetTextFontProperties />
        </div>
        <div v-if="filePageId == '02-36'">
          <SetTextMargins />
        </div>
        <div v-if="filePageId == '02-37'">
          <SetTextTransparency />
        </div>
        <div v-if="filePageId == '02-38'">
          <SuperscriptAndSubscript />
        </div>
        <!--03_Slides-->
        <!--Layout-->
        <div v-if="filePageId == '03-1-1'">
          <AddDifferentLayoutSlides />
        </div>
        <div v-if="filePageId == '03-1-2'">
          <ChangeSlideLayout />
        </div>
        <div v-if="filePageId == '03-1-3'">
          <GetSlideLayoutName />
        </div>
        <div v-if="filePageId == '03-1-4'">
          <SetSlideLayout />
        </div>
        <!--Transitions-->
        <div v-if="filePageId == '03-2-1'">
          <BetterSlideTransitions />
        </div>
        <div v-if="filePageId == '03-2-2'">
          <SetAdvanceAfterTime />
        </div>
        <div v-if="filePageId == '03-2-3'">
          <SetTransitionEffects />
        </div>
        <div v-if="filePageId == '03-2-4'">
          <SetTransitions />
        </div>

        <div v-if="filePageId == '03-3'">
          <AddImageInMaster />
        </div>
        <div v-if="filePageId == '03-4'">
          <AddSlideUsingMasterLayout />
        </div>
        <div v-if="filePageId == '03-5'">
          <AppendSlideWithMasterLayout />
        </div>
        <div v-if="filePageId == '03-6'">
          <ApplySlideMaster />
        </div>
        <div v-if="filePageId == '03-7'">
          <ChangeSlidePosition />
        </div>
        <div v-if="filePageId == '03-8'">
          <ClonePPTAtEndOfAnother />
        </div>
        <div v-if="filePageId == '03-9'">
          <ClonePPTMasterToAnother />
        </div>
        <div v-if="filePageId == '03-10'">
          <CloneSlideAtTheEnd />
        </div>
        <div v-if="filePageId == '03-11'">
          <CloneSlideToAnotherPPT />
        </div>
        <div v-if="filePageId == '03-12'">
          <CloneSlideWithinAPPT />
        </div>
        <div v-if="filePageId == '03-13'">
          <CreateSlide />
        </div>
        <div v-if="filePageId == '03-14'">
          <CreateSlideMasterAndApply />
        </div>
        <div v-if="filePageId == '03-15'">
          <DetectUsedThemes />
        </div>
        <div v-if="filePageId == '03-16'">
          <GetSlideByIndexOrID />
        </div>
        <div v-if="filePageId == '03-17'">
          <GetSlideText />
        </div>
        <div v-if="filePageId == '03-18'">
          <HideSlide />
        </div>
        <div v-if="filePageId == '03-19'">
          <MergeSelectedSlides />
        </div>
        <div v-if="filePageId == '03-20'">
          <RemoveSlide />
        </div>
        <div v-if="filePageId == '03-21'">
          <RemoveUnusedLayoutMaster />
        </div>
        <div v-if="filePageId == '03-22'">
          <SetStartingNumberForSlides />
        </div>
        <div v-if="filePageId == '03-23'">
          <SlideTitle />
        </div>
        <!--04_Shapes -->
        <div v-if="filePageId == '4-1'">
          <AddLineToSlide />
        </div>
        <div v-if="filePageId == '4-2'">
          <AddLineWithArrow />
        </div>
        <div v-if="filePageId == '4-3'">
          <AddLineWithTwoPoints />
        </div>
        <div v-if="filePageId == '4-4'">
          <AddRoundCornerRectangle />
        </div>
        <div v-if="filePageId == '4-5'">
          <AddShapes />
        </div>
        <div v-if="filePageId == '4-6'">
          <ArrangeShapes />
        </div>
        <div v-if="filePageId == '4-7'">
          <Background />
        </div>
        <div v-if="filePageId == '4-8'">
          <CopyShapesBetweenSlides />
        </div>
        <div v-if="filePageId == '4-9'">
          <FillShapeWithGradient />
        </div>
        <div v-if="filePageId == '4-10'">
          <FillShapeWithPattern />
        </div>
        <div v-if="filePageId == '4-11'">
          <FillShapeWithPicture />
        </div>
        <div v-if="filePageId == '4-12'">
          <FillShapeWithSolidColor />
        </div>
        <div v-if="filePageId == '4-13'">
          <FindShapeByAltText />
        </div>
        <div v-if="filePageId == '4-14'">
          <GetAllTitles />
        </div>
        <div v-if="filePageId == '4-15'">
          <GetShapeGroupAltText />
        </div>
        <div v-if="filePageId == '4-16'">
          <GetShapesByPlaceholder />
        </div>
        <div v-if="filePageId == '4-17'">
          <GroupShapes />
        </div>
        <div v-if="filePageId == '4-18'">
          <HideShape />
        </div>
        <div v-if="filePageId == '4-19'">
          <IsTextBox />
        </div>
        <div v-if="filePageId == '4-20'">
          <OperatePlaceholders />
        </div>
        <div v-if="filePageId == '4-21'">
          <PreventOrAllowChangingShape />
        </div>
        <div v-if="filePageId == '4-22'">
          <RemoveShape />
        </div>
        <div v-if="filePageId == '4-23'">
          <ReorderOverlappingShapes />
        </div>
        <div v-if="filePageId == '4-24'">
          <ResetPositionOfPlaceholder />
        </div>
        <div v-if="filePageId == '4-25'">
          <ResetShapeSizeAndPosition />
        </div>
        <div v-if="filePageId == '4-26'">
          <RotateShape />
        </div>
        <div v-if="filePageId == '4-27'">
          <Set3DEffectForShape />
        </div>
        <div v-if="filePageId == '4-28'">
          <SetAndGetAlternativeText />
        </div>
        <div v-if="filePageId == '4-29'">
          <SetEllipseFormat />
        </div>
        <div v-if="filePageId == '4-30'">
          <SetFormatForLines />
        </div>
        <div v-if="filePageId == '4-31'">
          <SetLineJoinStyles />
        </div>
        <div v-if="filePageId == '4-32'">
          <SetOutlineAndEffectForShape />
        </div>
        <div v-if="filePageId == '4-33'">
          <SetRadiusForRoundedRectangle />
        </div>
        <div v-if="filePageId == '4-34'">
          <SetRadiusOfRoundedRectangle />
        </div>
        <div v-if="filePageId == '4-35'">
          <SetRectangleFormat />
        </div>
        <div v-if="filePageId == '4-36'">
          <SetShadowEffectForShape />
        </div>
        <div v-if="filePageId == '4-37'">
          <ShapeToImage />
        </div>
        <div v-if="filePageId == '4-38'">
          <UngroupShapes />
        </div>
        <div v-if="filePageId == '4-39-1'">
          <AddExitAnimationForShape />
        </div>
        <div v-if="filePageId == '4-39-2'">
          <Animations />
        </div>
        <div v-if="filePageId == '4-39-3'">
          <ApplyAnimationOnChart />
        </div>
        <div v-if="filePageId == '4-39-4'">
          <ApplyAnimationOnShape />
        </div>
        <div v-if="filePageId == '4-39-5'">
          <ApplyAnimationOnText />
        </div>
        <div v-if="filePageId == '4-39-6'">
          <CustomPathAnimation />
        </div>
        <div v-if="filePageId == '4-39-7'">
          <DurationAndDelayTime />
        </div>
        <div v-if="filePageId == '4-39-8'">
          <GetAnimationEffectInfo />
        </div>
        <div v-if="filePageId == '4-39-9'">
          <GetAnimationsMotionPath />
        </div>
        <div v-if="filePageId == '4-39-10'">
          <SetAnimationForAnimateText />
        </div>
        <div v-if="filePageId == '4-39-11'">
          <SetAnimationRepeatType />
        </div>
        <!--05_DocumentOperation -->
        <div v-if="filePageId == '5-1'">
          <AddSection />
        </div>
        <div v-if="filePageId == '5-2'">
          <AddSlidetoSection />
        </div>
        <div v-if="filePageId == '5-3'">
          <DeleteSection />
        </div>
        <div v-if="filePageId == '5-4'">
          <GetSectionIndex />
        </div>
        <div v-if="filePageId == '5-5'">
          <LoadFromStream />
        </div>
        <div v-if="filePageId == '5-6'">
          <LoopPresentPPT />
        </div>
        <div v-if="filePageId == '5-7'">
          <PageSetup />
        </div>
        <div v-if="filePageId == '5-8'">
          <SaveToStream />
        </div>
        <div v-if="filePageId == '5-9'">
          <SetShowTypeAsKiosk />
        </div>
        <div v-if="filePageId == '5-10'">
          <SplitPPT />
        </div>
        <!--06_Properties -->
        <div v-if="filePageId == '6-1'">
          <GetBuiltinProperties />
        </div>
        <div v-if="filePageId == '6-2'">
          <MarkAsFinal />
        </div>
        <div v-if="filePageId == '6-3'">
          <Properties />
        </div>
        <div v-if="filePageId == '6-4'">
          <SetPropertiesForTemplate />
        </div>
        <!--07_Security -->
        <div v-if="filePageId == '7-1'">
          <CheckPasswordProtection />
        </div>
        <div v-if="filePageId == '7-2'">
          <Encrypt />
        </div>
        <div v-if="filePageId == '7-3'">
          <ModifyPasswordOfEncryptedPPT />
        </div>
        <div v-if="filePageId == '7-4'">
          <OpenEncryptedPPT />
        </div>
        <div v-if="filePageId == '7-5'">
          <RemoveAllDigitalSignatures />
        </div>
        <div v-if="filePageId == '7-6'">
          <RemoveEncryption />
        </div>
        <div v-if="filePageId == '7-7'">
          <SetDocumentReadOnly />
        </div>
        <!--08_Background -->
        <div v-if="filePageId == '8-1'">
          <SetBackground />
        </div>
        <div v-if="filePageId == '8-2'">
          <SetGradientBackground />
        </div>
        <div v-if="filePageId == '8-3'">
          <SetMasterBackground />
        </div>
        <!--09_Charts-->
        <div v-if="filePageId == '9-1'">
          <AddAndFormatErrorBars />
        </div>
        <div v-if="filePageId == '9-2'">
          <AddCustomErrorBars />
        </div>
        <div v-if="filePageId == '9-3'">
          <AddSecondaryValueAxis />
        </div>
        <div v-if="filePageId == '9-4'">
          <AddShadowEffectForDataLabel />
        </div>
        <div v-if="filePageId == '9-5'">
          <AddTrendLineForChartSeries />
        </div>
        <div v-if="filePageId == '9-6'">
          <AutoVaryColorForPieChart />
        </div>
        <div v-if="filePageId == '9-7'">
          <ChangeColorForLegend />
        </div>
        <div v-if="filePageId == '9-8'">
          <ChangeFontForChartDataTable />
        </div>
        <div v-if="filePageId == '9-9'">
          <ChangeFontSizeForLegend />
        </div>
        <div v-if="filePageId == '9-10'">
          <ChangeSeriesName />
        </div>
        <div v-if="filePageId == '9-11'">
          <ChangesForTrendLineEquation />
        </div>
        <div v-if="filePageId == '9-12'">
          <ChangeTextFontInChart />
        </div>
        <div v-if="filePageId == '9-13'">
          <ChartAxis />
        </div>
        <div v-if="filePageId == '9-14'">
          <CopyChartBetweenPptFiles />
        </div>
        <div v-if="filePageId == '9-15'">
          <CopyChartWithinOnePPT />
        </div>
        <div v-if="filePageId == '9-16'">
          <Create100PercentStackedBarChart />
        </div>
        <div v-if="filePageId == '9-17'">
          <CreateBoxAndWhiskerChart />
        </div>
        <div v-if="filePageId == '9-18'">
          <CreateBubbleChart />
        </div>
        <div v-if="filePageId == '9-19'">
          <CreateClusteredColumnChart />
        </div>
        <div v-if="filePageId == '9-20'">
          <CreateCombinationChart />
        </div>
        <div v-if="filePageId == '9-21'">
          <CreateCylinder3DClusteredChart />
        </div>
        <div v-if="filePageId == '9-22'">
          <CreateDoughnutChart />
        </div>
        <div v-if="filePageId == '9-23'">
          <CreateFunnelChart />
        </div>
        <div v-if="filePageId == '9-24'">
          <CreateHistogramChart />
        </div>
        <div v-if="filePageId == '9-25'">
          <CreateLineMarkersChart />
        </div>
        <div v-if="filePageId == '9-26'">
          <CreateMapChart />
        </div>
        <div v-if="filePageId == '9-27'">
          <CreateParetoChart />
        </div>
        <div v-if="filePageId == '9-28'">
          <CreatePieChart />
        </div>
        <div v-if="filePageId == '9-29'">
          <CreateScatterChart />
        </div>
        <div v-if="filePageId == '9-30'">
          <CreateSunBurstChart />
        </div>
        <div v-if="filePageId == '9-31'">
          <CreateTreeMapChart />
        </div>
        <div v-if="filePageId == '9-32'">
          <CreateWaterFallChart />
        </div>
        <div v-if="filePageId == '9-33'">
          <DeleteChartLegendEntries />
        </div>
        <div v-if="filePageId == '9-34'">
          <DetectIsSwitchRowAndColumn />
        </div>
        <div v-if="filePageId == '9-35'">
          <DoughnutChartHole />
        </div>
        <div v-if="filePageId == '9-36'">
          <EditChartData />
        </div>
        <div v-if="filePageId == '9-37'">
          <ExplodePieChart />
        </div>
        <div v-if="filePageId == '9-38'">
          <FillPictureInChartMarker />
        </div>
        <div v-if="filePageId == '9-39'">
          <FormatChartDataLabels />
        </div>
        <div v-if="filePageId == '9-40'">
          <GetValuesAndUnitFromAxis />
        </div>
        <div v-if="filePageId == '9-41'">
          <GroupTwoLevelAxisLabels />
        </div>
        <div v-if="filePageId == '9-42'">
          <HideAxisAndGridLine />
        </div>
        <div v-if="filePageId == '9-43'">
          <HideOrShowASeriesOfChart />
        </div>
        <div v-if="filePageId == '9-44'">
          <InvertIfNegative />
        </div>
        <div v-if="filePageId == '9-45'">
          <ModifyChartCategoryAxis />
        </div>
        <div v-if="filePageId == '9-46'">
          <MultipleCategoryChart />
        </div>
        <div v-if="filePageId == '9-47'">
          <ProtectChart />
        </div>
        <div v-if="filePageId == '9-48'">
          <RemoveChartFromPptSlide />
        </div>
        <div v-if="filePageId == '9-49'">
          <RemoveTickMarksOfAxis />
        </div>
        <div v-if="filePageId == '9-50'">
          <SaveChartAsImage />
        </div>
        <div v-if="filePageId == '9-51'">
          <ScaleBubbleChartSize />
        </div>
        <div v-if="filePageId == '9-52'">
          <SetAxisPosition />
        </div>
        <div v-if="filePageId == '9-53'">
          <SetAxisType />
        </div>
        <div v-if="filePageId == '9-54'">
          <SetChartBorderstyle />
        </div>
        <div v-if="filePageId == '9-55'">
          <SetChartDataLabelRange />
        </div>
        <div v-if="filePageId == '9-56'">
          <SetChartDataNumberFormat />
        </div>
        <div v-if="filePageId == '9-57'">
          <SetColorAndNameForTrendline />
        </div>
        <div v-if="filePageId == '9-58'">
          <SetDatalabelPosition />
        </div>
        <div v-if="filePageId == '9-59'">
          <SetDatapointColorInChart />
        </div>
        <div v-if="filePageId == '9-60'">
          <SetDisplayUnit />
        </div>
        <div v-if="filePageId == '9-61'">
          <SetDistanceFromAxis />
        </div>
        <div v-if="filePageId == '9-62'">
          <SetGapWidth />
        </div>
        <div v-if="filePageId == '9-63'">
          <SetLegendOptions />
        </div>
        <div v-if="filePageId == '9-64'">
          <SetNumberFormatForAxis />
        </div>
        <div v-if="filePageId == '9-65'">
          <SetPercentageForLabels />
        </div>
        <div v-if="filePageId == '9-66'">
          <SetPositionOfChartDataLabels />
        </div>
        <div v-if="filePageId == '9-67'">
          <SetRotationForChartTitle />
        </div>
        <div v-if="filePageId == '9-68'">
          <SetRotationForDataLabel />
        </div>
        <div v-if="filePageId == '9-69'">
          <SetRotationForValueAxisText />
        </div>
        <div v-if="filePageId == '9-70'">
          <SetSeriesLineColor />
        </div>
        <div v-if="filePageId == '9-71'">
          <SetSeriesOverlap />
        </div>
        <div v-if="filePageId == '9-72'">
          <SetSizeAndStyleForMarker />
        </div>
        <div v-if="filePageId == '9-73'">
          <SetSizeForChartPlotArea />
        </div>
        <div v-if="filePageId == '9-74'">
          <SetTextFontForChartTitle />
        </div>
        <div v-if="filePageId == '9-75'">
          <SetTextFontForLegendAndAxis />
        </div>
        <div v-if="filePageId == '9-76'">
          <SetTickMarkLabelsOnCategoryAxis />
        </div>
        <div v-if="filePageId == '9-77'">
          <SetTickMarksInterval />
        </div>
        <div v-if="filePageId == '9-78'">
          <ShowChartLabels />
        </div>
        <div v-if="filePageId == '9-79'">
          <VaryColorOfSameSerieDataMarker />
        </div>
        <!--10_Conversion-->
        <div v-if="filePageId == '10-1'">
          <ConvertODPtoPDF />
        </div>
        <div v-if="filePageId == '10-2'">
          <ConvertPdfWithDefaultFont />
        </div>
        <div v-if="filePageId == '10-3'">
          <ConvertPPSToPPTX />
        </div>
        <div v-if="filePageId == '10-4'">
          <ConvertPPTToOFD />
        </div>
        <div v-if="filePageId == '10-5'">
          <ConvertUnhideSlidesToPdf />
        </div>
        <div v-if="filePageId == '10-6'">
          <CustomSizeToTIFF />
        </div>
        <div v-if="filePageId == '10-7'">
          <IndividualSlideToHtml />
        </div>
        <div v-if="filePageId == '10-8'">
          <LoadSaveDPSAndDPT />
        </div>
        <div v-if="filePageId == '10-9'">
          <OdpToPdf />
        </div>
        <div v-if="filePageId == '10-10'">
          <OneSlideToSVG />
        </div>
        <div v-if="filePageId == '10-11'">
          <SlideToSVG />
        </div>
        <div v-if="filePageId == '10-12'">
          <SpecificSlideToPDF />
        </div>
        <div v-if="filePageId == '10-13'">
          <ToHTML />
        </div>
        <div v-if="filePageId == '10-14'">
          <ToImage />
        </div>
        <div v-if="filePageId == '10-15'">
          <ToMarkDown />
        </div>
        <div v-if="filePageId == '10-16'">
          <ToPDF />
        </div>
        <div v-if="filePageId == '10-17'">
          <ToPdfWithSpecificPageSize />
        </div>
        <div v-if="filePageId == '10-18'">
          <ToPPTX />
        </div>
        <div v-if="filePageId == '10-19'">
          <ToSpecificSizeImage />
        </div>
        <div v-if="filePageId == '10-20'">
          <ToSVG />
        </div>
        <div v-if="filePageId == '10-21'">
          <ToXPS />
        </div>

        <!--11_Images-->
        <div v-if="filePageId == '11-1'">
          <ChangeImageSize />
        </div>
        <div v-if="filePageId == '11-2'">
          <CropImage />
        </div>
        <div v-if="filePageId == '11-3'">
          <ExtractImage />
        </div>
        <div v-if="filePageId == '11-4'">
          <ExtractImageFromSpecificSlide />
        </div>
        <div v-if="filePageId == '11-5'">
          <InsertEMFInPPT />
        </div>
        <div v-if="filePageId == '11-6'">
          <InsertImage />
        </div>
        <div v-if="filePageId == '11-7'">
          <RemoveImages />
        </div>
        <div v-if="filePageId == '11-8'">
          <SetImageFrameFormat />
        </div>
        <div v-if="filePageId == '11-9'">
          <SetImageTransparency />
        </div>
        <div v-if="filePageId == '11-10'">
          <UpdateImage />
        </div>

        <!--12_Tables-->
        <div v-if="filePageId == '12-1'">
          <AddImageInTableCell />
        </div>
        <div v-if="filePageId == '12-2'">
          <AddRowToTable />
        </div>
        <div v-if="filePageId == '12-3'">
          <AdjustColumnByTextWidth />
        </div>
        <div v-if="filePageId == '12-4'">
          <CloneRowAndColumn />
        </div>
        <div v-if="filePageId == '12-5'">
          <CreateTable />
        </div>
        <div v-if="filePageId == '12-6'">
          <EditTableDataAndStyle />
        </div>
        <div v-if="filePageId == '12-7'">
          <FillAllCellsWithColor />
        </div>
        <div v-if="filePageId == '12-8'">
          <FillParticularRowWithColor />
        </div>
        <div v-if="filePageId == '12-9'">
          <GetBorderColorOfCell />
        </div>
        <div v-if="filePageId == '12-10'">
          <IdentifyMergedCells />
        </div>
        <div v-if="filePageId == '12-11'">
          <LockAspectRatio />
        </div>
        <!--12_Tables-->
        <div v-if="filePageId == '12-11'">
          <LockAspectRatio />
        </div>
        <div v-if="filePageId == '12-12'">
          <MergeTableCell />
        </div>
        <div v-if="filePageId == '12-13'">
          <RemoveRowsAndColumns />
        </div>
        <div v-if="filePageId == '12-14'">
          <RemoveTableBorderStyle />
        </div>
        <div v-if="filePageId == '12-15'">
          <RemoveTableFromPptSlide />
        </div>
        <div v-if="filePageId == '12-16'">
          <SetAlignmentInTable />
        </div>
        <div v-if="filePageId == '12-17'">
          <SetBordersForExistingTable />
        </div>
        <div v-if="filePageId == '12-18'">
          <SetBordersForNewlyTables />
        </div>
        <div v-if="filePageId == '12-19'">
          <SetFirstRowAsHeader />
        </div>
        <div v-if="filePageId == '12-20'">
          <SetRowHeightColumnWidth />
        </div>
        <div v-if="filePageId == '12-21'">
          <SetTableBorderStyle />
        </div>
        <div v-if="filePageId == '12-22'">
          <SetTableStyle />
        </div>
        <div v-if="filePageId == '12-23'">
          <SetTextFormat />
        </div>
        <div v-if="filePageId == '12-24'">
          <SplitSpecificTableCell />
        </div>
        <div v-if="filePageId == '12-25'">
          <TraverseThroughCells />
        </div>
        <!--13_Hyperlink-->
        <div v-if="filePageId == '13-1'">
          <AddHyperlinkToImage />
        </div>
        <div v-if="filePageId == '13-2'">
          <AddHyperlinkToSmartArtNode />
        </div>
        <div v-if="filePageId == '13-3'">
          <AddHyperlinkToText />
        </div>
        <div v-if="filePageId == '13-4'">
          <ChangeHyperlinkColor />
        </div>
        <div v-if="filePageId == '13-5'">
          <GetLinkedSlide />
        </div>
        <div v-if="filePageId == '13-6'">
          <HyperlinkOutlineStyle />
        </div>
        <div v-if="filePageId == '13-7'">
          <Hyperlinks />
        </div>
        <div v-if="filePageId == '13-8'">
          <LinkToASpecificSlide />
        </div>
        <div v-if="filePageId == '13-9'">
          <LinkToLastViewedSlide />
        </div>
        <div v-if="filePageId == '13-10'">
          <ModifyHyperlink />
        </div>
        <div v-if="filePageId == '13-11'">
          <RemoveHyperlink />
        </div>
        <!--14_AudioAndVideo-->
        <div v-if="filePageId == '14-1'">
          <ExtractAudio />
        </div>
        <div v-if="filePageId == '14-2'">
          <ExtractVideo />
        </div>
        <div v-if="filePageId == '14-3'">
          <HideAudioDuringShow />
        </div>
        <div v-if="filePageId == '14-4'">
          <InsertAudio />
        </div>
        <div v-if="filePageId == '14-5'">
          <InsertVideo />
        </div>
        <div v-if="filePageId == '14-6'">
          <ObtainSoundEffect />
        </div>
        <div v-if="filePageId == '14-7'">
          <ReplaceVideo />
        </div>
        <div v-if="filePageId == '14-8'">
          <SetPlayModeForVideo />
        </div>
        <!--15_CommentAndNote-->
        <div v-if="filePageId == '15-1'">
          <AddAndGetSpeakerNotes />
        </div>
        <div v-if="filePageId == '15-2'">
          <AddComment />
        </div>
        <div v-if="filePageId == '15-3'">
          <AddNote />
        </div>
        <div v-if="filePageId == '15-4'">
          <DeleteComment />
        </div>
        <div v-if="filePageId == '15-5'">
          <ExtractComments />
        </div>
        <div v-if="filePageId == '15-6'">
          <GetSlideComments />
        </div>
        <div v-if="filePageId == '15-7'">
          <PptToSvgRetainNotes />
        </div>
        <div v-if="filePageId == '15-8'">
          <RemoveNoteAtSpecificSlide />
        </div>
        <div v-if="filePageId == '15-9'">
          <RemoveSpeakerNotes />
        </div>
        <div v-if="filePageId == '15-10'">
          <ReplyToComment />
        </div>
        <!--16_HeaderAndFooter-->
        <div v-if="filePageId == '16-1'">
          <HeaderAndFooter />
        </div>
        <div v-if="filePageId == '16-2'">
          <ManageNoteMasterHeaderFooter />
        </div>
        <!--17_SmartArt-->
        <div v-if="filePageId == '17-1'">
          <AccessChildNodes />
        </div>
        <div v-if="filePageId == '17-2'">
          <AccessSmartArt />
        </div>
        <div v-if="filePageId == '17-3'">
          <AccessSmartArtLayout />
        </div>
        <div v-if="filePageId == '17-4'">
          <AccessSpecificChildNode />
        </div>
        <div v-if="filePageId == '17-5'">
          <AddNodeByPosition />
        </div>
        <div v-if="filePageId == '17-6'">
          <AddSmartArtNode />
        </div>
        <div v-if="filePageId == '17-7'">
          <AssistantNode />
        </div>
        <div v-if="filePageId == '17-8'">
          <ChangeNodeText />
        </div>
        <div v-if="filePageId == '17-9'">
          <ChangeSmartArtColorStyle />
        </div>
        <div v-if="filePageId == '17-10'">
          <ChangeSmartArtShapeStyle />
        </div>
        <div v-if="filePageId == '17-11'">
          <CreateSmartArtShape />
        </div>
        <div v-if="filePageId == '17-12'">
          <ExtractTextFromSmartArt />
        </div>
        <div v-if="filePageId == '17-13'">
          <RemoveNode />
        </div>
        <div v-if="filePageId == '17-14'">
          <SetSmartArtLinklineOutline />
        </div>
        <div v-if="filePageId == '17-15'">
          <SetSmartArtNodeOutline />
        </div>
        <!--18_Watermark -->
        <div v-if="filePageId == '18-1'">
          <AddImageWatermark />
        </div>
        <div v-if="filePageId == '18-2'">
          <AddWatermark />
        </div>
        <div v-if="filePageId == '18-3'">
          <RemoveTextOrImageWatermark />
        </div>
        <!--19_OLE -->
        <div v-if="filePageId == '19-1'">
          <EmbedExcelAsOLE />
        </div>
        <div v-if="filePageId == '19-2'">
          <EmbedZipIntoPPT />
        </div>
        <div v-if="filePageId == '19-3'">
          <ExtractOLEObject />
        </div>
        <div v-if="filePageId == '19-4'">
          <ModifyOLEData />
        </div>
        <!--20_VBA -->
        <div v-if="filePageId == '20-1'">
          <RemoveVBAMacros />
        </div>

      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref } from "vue";
import SpireTree from "@/components/SpireTree.vue";
import Description from "@/components/Description.vue";

//01_QuickGuide
import HelloWorld from "@/components/01_QuickGuide/HelloWorld.vue";
//02_ParagraphAndText
import AddParagraph from "@/components/02_ParagraphAndText/AddParagraph.vue";
import Alignment from "@/components/02_ParagraphAndText/Alignment.vue";
import AppendHTML from "@/components/02_ParagraphAndText/AppendHTML.vue";
import AutoFitTextOrShape from "@/components/02_ParagraphAndText/AutoFitTextOrShape.vue";
import BordersAndShading from "@/components/02_ParagraphAndText/BordersAndShading.vue";
import Bullets from "@/components/02_ParagraphAndText/Bullets.vue";
import ChangeTextStyle from "@/components/02_ParagraphAndText/ChangeTextStyle.vue";
import CopyParagraphToAnotherPPT from "@/components/02_ParagraphAndText/CopyParagraphToAnotherPPT.vue";
import CustomBulletsNumber from "@/components/02_ParagraphAndText/CustomBulletsNumber.vue";
import EditPromptText from "@/components/02_ParagraphAndText/EditPromptText.vue";
import ExtractText from "@/components/02_ParagraphAndText/ExtractText.vue";
import GetTextFrameEffectiveData from "@/components/02_ParagraphAndText/GetTextFrameEffectiveData.vue";
import GetTextStyleEffectiveData from "@/components/02_ParagraphAndText/GetTextStyleEffectiveData.vue";
import HighlightSpecifiedText from "@/components/02_ParagraphAndText/HighlightSpecifiedText.vue";
import Indent from "@/components/02_ParagraphAndText/Indent.vue";
import LineSpacing from "@/components/02_ParagraphAndText/LineSpacing.vue";
import MixFontStyles from "@/components/02_ParagraphAndText/MixFontStyles.vue";
import ModifyStyleOfFirstFoundText from "@/components/02_ParagraphAndText/ModifyStyleOfFirstFoundText.vue";
import MultipleLevelBullets from "@/components/02_ParagraphAndText/MultipleLevelBullets.vue";
import MultipleParagraphs from "@/components/02_ParagraphAndText/MultipleParagraphs.vue";
import PictureCustomBulletStyle from "@/components/02_ParagraphAndText/PictureCustomBulletStyle.vue";
import RemoveTextBox from "@/components/02_ParagraphAndText/RemoveTextBox.vue";
import ReplaceText from "@/components/02_ParagraphAndText/ReplaceText.vue";
import ReplaceTextRetentionStyle from "@/components/02_ParagraphAndText/ReplaceTextRetentionStyle.vue";
import ReplaceTextWithRegex from "@/components/02_ParagraphAndText/ReplaceTextWithRegex.vue";
import RotateText from "@/components/02_ParagraphAndText/RotateText.vue";
import Set3DEffectForText from "@/components/02_ParagraphAndText/Set3DEffectForText.vue";
import SetAnchorOfTextFrame from "@/components/02_ParagraphAndText/SetAnchorOfTextFrame.vue";
import SetColumnsCountOfTextFrame from "@/components/02_ParagraphAndText/SetColumnsCountOfTextFrame.vue";
import SetCustomFonts from "@/components/02_ParagraphAndText/SetCustomFonts.vue";
import SetParagraphFont from "@/components/02_ParagraphAndText/SetParagraphFont.vue";
import SetRightToLeftColumns from "@/components/02_ParagraphAndText/SetRightToLeftColumns.vue";
import SetShadowEffectForText from "@/components/02_ParagraphAndText/SetShadowEffectForText.vue";
import SetTextDirection from "@/components/02_ParagraphAndText/SetTextDirection.vue";
import SetTextFontProperties from "@/components/02_ParagraphAndText/SetTextFontProperties.vue";
import SetTextMargins from "@/components/02_ParagraphAndText/SetTextMargins.vue";
import SetTextTransparency from "@/components/02_ParagraphAndText/SetTextTransparency.vue";
import SuperscriptAndSubscript from "@/components/02_ParagraphAndText/SuperscriptAndSubscript.vue";
//03_Slides
import AddDifferentLayoutSlides from "@/components/03_Slides/Layout/AddDifferentLayoutSlides.vue";
import ChangeSlideLayout from "@/components/03_Slides/Layout/ChangeSlideLayout.vue";
import GetSlideLayoutName from "@/components/03_Slides/Layout/GetSlideLayoutName.vue";
import SetSlideLayout from "@/components/03_Slides/Layout/SetSlideLayout.vue";
import BetterSlideTransitions from "@/components/03_Slides/Transitions/BetterSlideTransitions.vue";
import SetAdvanceAfterTime from "@/components/03_Slides/Transitions/SetAdvanceAfterTime.vue";
import SetTransitionEffects from "@/components/03_Slides/Transitions/SetTransitionEffects.vue";
import SetTransitions from "@/components/03_Slides/Transitions/SetTransitions.vue";
import AddImageInMaster from "@/components/03_Slides/AddImageInMaster.vue";
import AddSlideUsingMasterLayout from "@/components/03_Slides/AddSlideUsingMasterLayout.vue";
import AppendSlideWithMasterLayout from "@/components/03_Slides/AppendSlideWithMasterLayout.vue";
import ApplySlideMaster from "@/components/03_Slides/ApplySlideMaster.vue";
import ChangeSlidePosition from "@/components/03_Slides/ChangeSlidePosition.vue";
import ClonePPTAtEndOfAnother from "@/components/03_Slides/ClonePPTAtEndOfAnother.vue";
import ClonePPTMasterToAnother from "@/components/03_Slides/ClonePPTMasterToAnother.vue";
import CloneSlideAtTheEnd from "@/components/03_Slides/CloneSlideAtTheEnd.vue";
import CloneSlideToAnotherPPT from "@/components/03_Slides/CloneSlideToAnotherPPT.vue";
import CloneSlideWithinAPPT from "@/components/03_Slides/CloneSlideWithinAPPT.vue";
import CreateSlide from "@/components/03_Slides/CreateSlide.vue";
import CreateSlideMasterAndApply from "@/components/03_Slides/CreateSlideMasterAndApply.vue";
import DetectUsedThemes from "@/components/03_Slides/DetectUsedThemes.vue";
import GetSlideByIndexOrID from "@/components/03_Slides/GetSlideByIndexOrID.vue";
import GetSlideText from "@/components/03_Slides/GetSlideText.vue";
import HideSlide from "@/components/03_Slides/HideSlide.vue";
import MergeSelectedSlides from "@/components/03_Slides/MergeSelectedSlides.vue";
import RemoveSlide from "@/components/03_Slides/RemoveSlide.vue";
import RemoveUnusedLayoutMaster from "@/components/03_Slides/RemoveUnusedLayoutMaster.vue";
import SetStartingNumberForSlides from "@/components/03_Slides/SetStartingNumberForSlides.vue";
import SlideTitle from "@/components/03_Slides/SlideTitle.vue";

//04_Shapes
import AddLineToSlide from "@/components/04_Shapes/AddLineToSlide.vue";
import AddLineWithArrow from "@/components/04_Shapes/AddLineWithArrow.vue";
import AddLineWithTwoPoints from "@/components/04_Shapes/AddLineWithTwoPoints.vue";
import AddRoundCornerRectangle from "@/components/04_Shapes/AddRoundCornerRectangle.vue";
import AddShapes from "@/components/04_Shapes/AddShapes.vue";
import ArrangeShapes from "@/components/04_Shapes/ArrangeShapes.vue";
import Background from "@/components/04_Shapes/Background.vue";
import CopyShapesBetweenSlides from "@/components/04_Shapes/CopyShapesBetweenSlides.vue";
import FillShapeWithGradient from "@/components/04_Shapes/FillShapeWithGradient.vue";
import FillShapeWithPattern from "@/components/04_Shapes/FillShapeWithPattern.vue";
import FillShapeWithPicture from "@/components/04_Shapes/FillShapeWithPicture.vue";
import FillShapeWithSolidColor from "@/components/04_Shapes/FillShapeWithSolidColor.vue";
import FindShapeByAltText from "@/components/04_Shapes/FindShapeByAltText.vue";
import GetAllTitles from "@/components/04_Shapes/GetAllTitles.vue";
import GetShapeGroupAltText from "@/components/04_Shapes/GetShapeGroupAltText.vue";
import GetShapesByPlaceholder from "@/components/04_Shapes/GetShapesByPlaceholder.vue";
import GroupShapes from "@/components/04_Shapes/GroupShapes.vue";
import HideShape from "@/components/04_Shapes/HideShape.vue";
import IsTextBox from "@/components/04_Shapes/IsTextBox.vue";
import OperatePlaceholders from "@/components/04_Shapes/OperatePlaceholders.vue";
import PreventOrAllowChangingShape from "@/components/04_Shapes/PreventOrAllowChangingShape.vue";
import RemoveShape from "@/components/04_Shapes/RemoveShape.vue";
import ReorderOverlappingShapes from "@/components/04_Shapes/ReorderOverlappingShapes.vue";
import ResetPositionOfPlaceholder from "@/components/04_Shapes/ResetPositionOfPlaceholder.vue";
import ResetShapeSizeAndPosition from "@/components/04_Shapes/ResetShapeSizeAndPosition.vue";
import RotateShape from "@/components/04_Shapes/RotateShape.vue";
import Set3DEffectForShape from "@/components/04_Shapes/Set3DEffectForShape.vue";
import SetAndGetAlternativeText from "@/components/04_Shapes/SetAndGetAlternativeText.vue";
import SetEllipseFormat from "@/components/04_Shapes/SetEllipseFormat.vue";
import SetFormatForLines from "@/components/04_Shapes/SetFormatForLines.vue";
import SetLineJoinStyles from "@/components/04_Shapes/SetLineJoinStyles.vue";
import SetOutlineAndEffectForShape from "@/components/04_Shapes/SetOutlineAndEffectForShape.vue";
import SetRadiusForRoundedRectangle from "@/components/04_Shapes/SetRadiusForRoundedRectangle.vue";
import SetRadiusOfRoundedRectangle from "@/components/04_Shapes/SetRadiusOfRoundedRectangle.vue";
import SetRectangleFormat from "@/components/04_Shapes/SetRectangleFormat.vue";
import SetShadowEffectForShape from "@/components/04_Shapes/SetShadowEffectForShape.vue";
import ShapeToImage from "@/components/04_Shapes/ShapeToImage.vue";
import UngroupShapes from "@/components/04_Shapes/UngroupShapes.vue";
import AddExitAnimationForShape from "@/components/04_Shapes/Animations/AddExitAnimationForShape.vue";
import Animations from "@/components/04_Shapes/Animations/Animations.vue";
import ApplyAnimationOnChart from "@/components/04_Shapes/Animations/ApplyAnimationOnChart.vue";
import ApplyAnimationOnShape from "@/components/04_Shapes/Animations/ApplyAnimationOnShape.vue";
import ApplyAnimationOnText from "@/components/04_Shapes/Animations/ApplyAnimationOnText.vue";
import CustomPathAnimation from "@/components/04_Shapes/Animations/CustomPathAnimation.vue";
import DurationAndDelayTime from "@/components/04_Shapes/Animations/DurationAndDelayTime.vue";
import GetAnimationEffectInfo from "@/components/04_Shapes/Animations/GetAnimationEffectInfo.vue";
import GetAnimationsMotionPath from "@/components/04_Shapes/Animations/GetAnimationsMotionPath.vue";
import SetAnimationForAnimateText from "@/components/04_Shapes/Animations/SetAnimationForAnimateText.vue";
import SetAnimationRepeatType from "@/components/04_Shapes/Animations/SetAnimationRepeatType.vue";

//05_DocumentOperation
import AddSection from "@/components/05_DocumentOperation/AddSection.vue";
import AddSlidetoSection from "@/components/05_DocumentOperation/AddSlidetoSection.vue";
import DeleteSection from "@/components/05_DocumentOperation/DeleteSection.vue";
import GetSectionIndex from "@/components/05_DocumentOperation/GetSectionIndex.vue";
import LoadFromStream from "@/components/05_DocumentOperation/LoadFromStream.vue";
import LoopPresentPPT from "@/components/05_DocumentOperation/LoopPresentPPT.vue";
import PageSetup from "@/components/05_DocumentOperation/PageSetup.vue";
import SaveToStream from "@/components/05_DocumentOperation/SaveToStream.vue";
import SetShowTypeAsKiosk from "@/components/05_DocumentOperation/SetShowTypeAsKiosk.vue";
import SplitPPT from "@/components/05_DocumentOperation/SplitPPT.vue";

//06_Properties
import GetBuiltinProperties from "@/components/06_Properties/GetBuiltinProperties.vue";
import MarkAsFinal from "@/components/06_Properties/MarkAsFinal.vue";
import Properties from "@/components/06_Properties/Properties.vue";
import SetPropertiesForTemplate from "@/components/06_Properties/SetPropertiesForTemplate.vue";

//07_Security
import CheckPasswordProtection from "@/components/07_Security/CheckPasswordProtection.vue";
import Encrypt from "@/components/07_Security/Encrypt.vue";
import ModifyPasswordOfEncryptedPPT from "@/components/07_Security/ModifyPasswordOfEncryptedPPT.vue";
import OpenEncryptedPPT from "@/components/07_Security/OpenEncryptedPPT.vue";
import RemoveAllDigitalSignatures from "@/components/07_Security/RemoveAllDigitalSignatures.vue";
import RemoveEncryption from "@/components/07_Security/RemoveEncryption.vue";
import SetDocumentReadOnly from "@/components/07_Security/SetDocumentReadOnly.vue";

//08_Background
import SetBackground from "@/components/08_Background/SetBackground.vue";
import SetGradientBackground from "@/components/08_Background/SetGradientBackground.vue";
import SetMasterBackground from "@/components/08_Background/SetMasterBackground.vue";

//09_Charts
import AddAndFormatErrorBars from "@/components/09_Charts/AddAndFormatErrorBars.vue";
import AddCustomErrorBars from "@/components/09_Charts/AddCustomErrorBars.vue";
import AddSecondaryValueAxis from "@/components/09_Charts/AddSecondaryValueAxis.vue";
import AddShadowEffectForDataLabel from "@/components/09_Charts/AddShadowEffectForDataLabel.vue";
import AddTrendLineForChartSeries from "@/components/09_Charts/AddTrendLineForChartSeries.vue";
import AutoVaryColorForPieChart from "@/components/09_Charts/AutoVaryColorForPieChart.vue";
import ChangeColorForLegend from "@/components/09_Charts/ChangeColorForLegend.vue";
import ChangeFontForChartDataTable from "@/components/09_Charts/ChangeFontForChartDataTable.vue";
import ChangeFontSizeForLegend from "@/components/09_Charts/ChangeFontSizeForLegend.vue";
import ChangeSeriesName from "@/components/09_Charts/ChangeSeriesName.vue";
import ChangesForTrendLineEquation from "@/components/09_Charts/ChangesForTrendLineEquation.vue";
import ChangeTextFontInChart from "@/components/09_Charts/ChangeTextFontInChart.vue";
import ChartAxis from "@/components/09_Charts/ChartAxis.vue";
import CopyChartBetweenPptFiles from "@/components/09_Charts/CopyChartBetweenPptFiles.vue";
import CopyChartWithinOnePPT from "@/components/09_Charts/CopyChartWithinOnePPT.vue";
import Create100PercentStackedBarChart from "@/components/09_Charts/Create100PercentStackedBarChart.vue";
import CreateBoxAndWhiskerChart from "@/components/09_Charts/CreateBoxAndWhiskerChart.vue";
import CreateBubbleChart from "@/components/09_Charts/CreateBubbleChart.vue";
import CreateClusteredColumnChart from "@/components/09_Charts/CreateClusteredColumnChart.vue";
import CreateCombinationChart from "@/components/09_Charts/CreateCombinationChart.vue";
import CreateCylinder3DClusteredChart from "@/components/09_Charts/CreateCylinder3DClusteredChart.vue";
import CreateDoughnutChart from "@/components/09_Charts/CreateDoughnutChart.vue";
import CreateFunnelChart from "@/components/09_Charts/CreateFunnelChart.vue";
import CreateHistogramChart from "@/components/09_Charts/CreateHistogramChart.vue";
import CreateLineMarkersChart from "@/components/09_Charts/CreateLineMarkersChart.vue";
import CreateMapChart from "@/components/09_Charts/CreateMapChart.vue";
import CreateParetoChart from "@/components/09_Charts/CreateParetoChart.vue";
import CreatePieChart from "@/components/09_Charts/CreatePieChart.vue";
import CreateScatterChart from "@/components/09_Charts/CreateScatterChart.vue";
import CreateSunBurstChart from "@/components/09_Charts/CreateSunBurstChart.vue";
import CreateTreeMapChart from "@/components/09_Charts/CreateTreeMapChart.vue";
import CreateWaterFallChart from "@/components/09_Charts/CreateWaterFallChart.vue";
import DeleteChartLegendEntries from "@/components/09_Charts/DeleteChartLegendEntries.vue";
import DetectIsSwitchRowAndColumn from "@/components/09_Charts/DetectIsSwitchRowAndColumn.vue";
import DoughnutChartHole from "@/components/09_Charts/DoughnutChartHole.vue";
import EditChartData from "@/components/09_Charts/EditChartData.vue";
import ExplodePieChart from "@/components/09_Charts/ExplodePieChart.vue";
import FillPictureInChartMarker from "@/components/09_Charts/FillPictureInChartMarker.vue";
import FormatChartDataLabels from "@/components/09_Charts/FormatChartDataLabels.vue";
import GetValuesAndUnitFromAxis from "@/components/09_Charts/GetValuesAndUnitFromAxis.vue";
import GroupTwoLevelAxisLabels from "@/components/09_Charts/GroupTwoLevelAxisLabels.vue";
import HideAxisAndGridLine from "@/components/09_Charts/HideAxisAndGridLine.vue";
import HideOrShowASeriesOfChart from "@/components/09_Charts/HideOrShowASeriesOfChart.vue";
import InvertIfNegative from "@/components/09_Charts/InvertIfNegative.vue";
import ModifyChartCategoryAxis from "@/components/09_Charts/ModifyChartCategoryAxis.vue";
import MultipleCategoryChart from "@/components/09_Charts/MultipleCategoryChart.vue";
import ProtectChart from "@/components/09_Charts/ProtectChart.vue";
import RemoveChartFromPptSlide from "@/components/09_Charts/RemoveChartFromPptSlide.vue";
import RemoveTickMarksOfAxis from "@/components/09_Charts/RemoveTickMarksOfAxis.vue";
import SaveChartAsImage from "@/components/09_Charts/SaveChartAsImage.vue";
import ScaleBubbleChartSize from "@/components/09_Charts/ScaleBubbleChartSize.vue";
import SetAxisPosition from "@/components/09_Charts/SetAxisPosition.vue";
import SetAxisType from "@/components/09_Charts/SetAxisType.vue";
import SetChartBorderstyle from "@/components/09_Charts/SetChartBorderstyle.vue";
import SetChartDataLabelRange from "@/components/09_Charts/SetChartDataLabelRange.vue";
import SetChartDataNumberFormat from "@/components/09_Charts/SetChartDataNumberFormat.vue";
import SetColorAndNameForTrendline from "@/components/09_Charts/SetColorAndNameForTrendline.vue";
import SetDatalabelPosition from "@/components/09_Charts/SetDatalabelPosition.vue";
import SetDatapointColorInChart from "@/components/09_Charts/SetDatapointColorInChart.vue";
import SetDisplayUnit from "@/components/09_Charts/SetDisplayUnit.vue";
import SetDistanceFromAxis from "@/components/09_Charts/SetDistanceFromAxis.vue";
import SetGapWidth from "@/components/09_Charts/SetGapWidth.vue";
import SetLegendOptions from "@/components/09_Charts/SetLegendOptions.vue";
import SetNumberFormatForAxis from "@/components/09_Charts/SetNumberFormatForAxis.vue";
import SetPercentageForLabels from "@/components/09_Charts/SetPercentageForLabels.vue";
import SetPositionOfChartDataLabels from "@/components/09_Charts/SetPositionOfChartDataLabels.vue";
import SetRotationForChartTitle from "@/components/09_Charts/SetRotationForChartTitle.vue";
import SetRotationForDataLabel from "@/components/09_Charts/SetRotationForDataLabel.vue";
import SetRotationForValueAxisText from "@/components/09_Charts/SetRotationForValueAxisText.vue";
import SetSeriesLineColor from "@/components/09_Charts/SetSeriesLineColor.vue";
import SetSeriesOverlap from "@/components/09_Charts/SetSeriesOverlap.vue";
import SetSizeAndStyleForMarker from "@/components/09_Charts/SetSizeAndStyleForMarker.vue";
import SetSizeForChartPlotArea from "@/components/09_Charts/SetSizeForChartPlotArea.vue";
import SetTextFontForChartTitle from "@/components/09_Charts/SetTextFontForChartTitle.vue";
import SetTextFontForLegendAndAxis from "@/components/09_Charts/SetTextFontForLegendAndAxis.vue";
import SetTickMarkLabelsOnCategoryAxis from "@/components/09_Charts/SetTickMarkLabelsOnCategoryAxis.vue";
import SetTickMarksInterval from "@/components/09_Charts/SetTickMarksInterval.vue";
import ShowChartLabels from "@/components/09_Charts/ShowChartLabels.vue";
import VaryColorOfSameSerieDataMarker from "@/components/09_Charts/VaryColorOfSameSerieDataMarker.vue";

//10_Conversion
import ConvertODPtoPDF from "@/components/10_Conversion/ConvertODPtoPDF.vue";
import ConvertPdfWithDefaultFont from "@/components/10_Conversion/ConvertPdfWithDefaultFont.vue";
import ConvertPPSToPPTX from "@/components/10_Conversion/ConvertPPSToPPTX.vue";
import ConvertPPTToOFD from "@/components/10_Conversion/ConvertPPTToOFD.vue";
import ConvertUnhideSlidesToPdf from "@/components/10_Conversion/ConvertUnhideSlidesToPdf.vue";
import CustomSizeToTIFF from "@/components/10_Conversion/CustomSizeToTIFF.vue";
import IndividualSlideToHtml from "@/components/10_Conversion/IndividualSlideToHtml.vue";
import LoadSaveDPSAndDPT from "@/components/10_Conversion/LoadSaveDPSAndDPT.vue";
import OdpToPdf from "@/components/10_Conversion/OdpToPdf.vue";
import OneSlideToSVG from "@/components/10_Conversion/OneSlideToSVG.vue";
import SlideToSVG from "@/components/10_Conversion/SlideToSVG.vue";
import SpecificSlideToPDF from "@/components/10_Conversion/SpecificSlideToPDF.vue";
import ToHTML from "@/components/10_Conversion/ToHTML.vue";
import ToImage from "@/components/10_Conversion/ToImage.vue";
import ToMarkDown from "@/components/10_Conversion/ToMarkDown.vue";
import ToPDF from "@/components/10_Conversion/ToPDF.vue";
import ToPdfWithSpecificPageSize from "@/components/10_Conversion/ToPdfWithSpecificPageSize.vue";
import ToPPTX from "@/components/10_Conversion/ToPPTX.vue";
import ToSpecificSizeImage from "@/components/10_Conversion/ToSpecificSizeImage.vue";
import ToSVG from "@/components/10_Conversion/ToSVG.vue";
import ToXPS from "@/components/10_Conversion/ToXPS.vue";

//11_Images
import ChangeImageSize from "@/components/11_Images/ChangeImageSize.vue";
import CropImage from "@/components/11_Images/CropImage.vue";
import ExtractImage from "@/components/11_Images/ExtractImage.vue";
import ExtractImageFromSpecificSlide from "@/components/11_Images/ExtractImageFromSpecificSlide.vue";
import InsertEMFInPPT from "@/components/11_Images/InsertEMFInPPT.vue";
import InsertImage from "@/components/11_Images/InsertImage.vue";
import RemoveImages from "@/components/11_Images/RemoveImages.vue";
import SetImageFrameFormat from "@/components/11_Images/SetImageFrameFormat.vue";
import SetImageTransparency from "@/components/11_Images/SetImageTransparency.vue";
import UpdateImage from "@/components/11_Images/UpdateImage.vue";

//12_Tables
import AddImageInTableCell from "@/components/12_Tables/AddImageInTableCell.vue";
import AddRowToTable from "@/components/12_Tables/AddRowToTable.vue";
import AdjustColumnByTextWidth from "@/components/12_Tables/AdjustColumnByTextWidth.vue";
import CloneRowAndColumn from "@/components/12_Tables/CloneRowAndColumn.vue";
import CreateTable from "@/components/12_Tables/CreateTable.vue";
import EditTableDataAndStyle from "@/components/12_Tables/EditTableDataAndStyle.vue";
import FillAllCellsWithColor from "@/components/12_Tables/FillAllCellsWithColor.vue";
import FillParticularRowWithColor from "@/components/12_Tables/FillParticularRowWithColor.vue";
import GetBorderColorOfCell from "@/components/12_Tables/GetBorderColorOfCell.vue";
import IdentifyMergedCells from "@/components/12_Tables/IdentifyMergedCells.vue";
import LockAspectRatio from "@/components/12_Tables/LockAspectRatio.vue";
import MergeTableCell from "@/components/12_Tables/MergeTableCell.vue";
import RemoveRowsAndColumns from "@/components/12_Tables/RemoveRowsAndColumns.vue";
import RemoveTableBorderStyle from "@/components/12_Tables/RemoveTableBorderStyle.vue";
import RemoveTableFromPptSlide from "@/components/12_Tables/RemoveTableFromPptSlide.vue";
import SetAlignmentInTable from "@/components/12_Tables/SetAlignmentInTable.vue";
import SetBordersForExistingTable from "@/components/12_Tables/SetBordersForExistingTable.vue";
import SetBordersForNewlyTables from "@/components/12_Tables/SetBordersForNewlyTables.vue";
import SetFirstRowAsHeader from "@/components/12_Tables/SetFirstRowAsHeader.vue";
import SetRowHeightColumnWidth from "@/components/12_Tables/SetRowHeightColumnWidth.vue";
import SetTableBorderStyle from "@/components/12_Tables/SetTableBorderStyle.vue";
import SetTableStyle from "@/components/12_Tables/SetTableStyle.vue";
import SetTextFormat from "@/components/12_Tables/SetTextFormat.vue";
import SplitSpecificTableCell from "@/components/12_Tables/SplitSpecificTableCell.vue";
import TraverseThroughCells from "@/components/12_Tables/TraverseThroughCells.vue";

//13_Hyperlink
import AddHyperlinkToImage from "@/components/13_Hyperlink/AddHyperlinkToImage.vue";
import AddHyperlinkToSmartArtNode from "@/components/13_Hyperlink/AddHyperlinkToSmartArtNode.vue";
import AddHyperlinkToText from "@/components/13_Hyperlink/AddHyperlinkToText.vue";
import ChangeHyperlinkColor from "@/components/13_Hyperlink/ChangeHyperlinkColor.vue";
import GetLinkedSlide from "@/components/13_Hyperlink/GetLinkedSlide.vue";
import HyperlinkOutlineStyle from "@/components/13_Hyperlink/HyperlinkOutlineStyle.vue";
import Hyperlinks from "@/components/13_Hyperlink/Hyperlinks.vue";
import LinkToASpecificSlide from "@/components/13_Hyperlink/LinkToASpecificSlide.vue";
import LinkToLastViewedSlide from "@/components/13_Hyperlink/LinkToLastViewedSlide.vue";
import ModifyHyperlink from "@/components/13_Hyperlink/ModifyHyperlink.vue";
import RemoveHyperlink from "@/components/13_Hyperlink/RemoveHyperlink.vue";

//14_AudioAndVideo
import ExtractAudio from "@/components/14_AudioAndVideo/ExtractAudio.vue";
import ExtractVideo from "@/components/14_AudioAndVideo/ExtractVideo.vue";
import HideAudioDuringShow from "@/components/14_AudioAndVideo/HideAudioDuringShow.vue";
import InsertAudio from "@/components/14_AudioAndVideo/InsertAudio.vue";
import InsertVideo from "@/components/14_AudioAndVideo/InsertVideo.vue";
import ObtainSoundEffect from "@/components/14_AudioAndVideo/ObtainSoundEffect.vue";
import ReplaceVideo from "@/components/14_AudioAndVideo/ReplaceVideo.vue";
import SetPlayModeForVideo from "@/components/14_AudioAndVideo/SetPlayModeForVideo.vue";

//15_CommentAndNote
import AddAndGetSpeakerNotes from "@/components/15_CommentAndNote/AddAndGetSpeakerNotes.vue";
import AddComment from "@/components/15_CommentAndNote/AddComment.vue";
import AddNote from "@/components/15_CommentAndNote/AddNote.vue";
import DeleteComment from "@/components/15_CommentAndNote/DeleteComment.vue";
import ExtractComments from "@/components/15_CommentAndNote/ExtractComments.vue";
import GetSlideComments from "@/components/15_CommentAndNote/GetSlideComments.vue";
import PptToSvgRetainNotes from "@/components/15_CommentAndNote/PptToSvgRetainNotes.vue";
import RemoveNoteAtSpecificSlide from "@/components/15_CommentAndNote/RemoveNoteAtSpecificSlide.vue";
import RemoveSpeakerNotes from "@/components/15_CommentAndNote/RemoveSpeakerNotes.vue";
import ReplyToComment from "@/components/15_CommentAndNote/ReplyToComment.vue";

//16_HeaderAndFooter
import HeaderAndFooter from "@/components/16_HeaderAndFooter/HeaderAndFooter.vue";
import ManageNoteMasterHeaderFooter from "@/components/16_HeaderAndFooter/ManageNoteMasterHeaderFooter.vue";

//17_SmartArt
import AccessChildNodes from "@/components/17_SmartArt/AccessChildNodes.vue";
import AccessSmartArt from "@/components/17_SmartArt/AccessSmartArt.vue";
import AccessSmartArtLayout from "@/components/17_SmartArt/AccessSmartArtLayout.vue";
import AccessSpecificChildNode from "@/components/17_SmartArt/AccessSpecificChildNode.vue";
import AddNodeByPosition from "@/components/17_SmartArt/AddNodeByPosition.vue";
import AddSmartArtNode from "@/components/17_SmartArt/AddSmartArtNode.vue";
import AssistantNode from "@/components/17_SmartArt/AssistantNode.vue";
import ChangeNodeText from "@/components/17_SmartArt/ChangeNodeText.vue";
import ChangeSmartArtColorStyle from "@/components/17_SmartArt/ChangeSmartArtColorStyle.vue";
import ChangeSmartArtShapeStyle from "@/components/17_SmartArt/ChangeSmartArtShapeStyle.vue";
import CreateSmartArtShape from "@/components/17_SmartArt/CreateSmartArtShape.vue";
import ExtractTextFromSmartArt from "@/components/17_SmartArt/ExtractTextFromSmartArt.vue";
import RemoveNode from "@/components/17_SmartArt/RemoveNode.vue";
import SetSmartArtLinklineOutline from "@/components/17_SmartArt/SetSmartArtLinklineOutline.vue";
import SetSmartArtNodeOutline from "@/components/17_SmartArt/SetSmartArtNodeOutline.vue";

//18_Watermark
import AddImageWatermark from "@/components/18_Watermark/AddImageWatermark.vue";
import AddWatermark from "@/components/18_Watermark/AddWatermark.vue";
import RemoveTextOrImageWatermark from "@/components/18_Watermark/RemoveTextOrImageWatermark.vue";

//19_OLE
import EmbedExcelAsOLE from "@/components/19_OLE/EmbedExcelAsOLE.vue";
import EmbedZipIntoPPT from "@/components/19_OLE/EmbedZipIntoPPT.vue";
import ExtractOLEObject from "@/components/19_OLE/ExtractOLEObject.vue";
import ModifyOLEData from "@/components/19_OLE/ModifyOLEData.vue";

//20_VBA
import RemoveVBAMacros from "@/components/20_VBA/RemoveVBAMacros.vue";



const filePageId = ref<string>("0");
const handleChildID = (id: any) => {
  filePageId.value = id;
};
</script>

<style lang="scss" scoped>
* {
  padding: 0%;
  margin: 0%;
}

button {
  padding: 5px 10px;
}

.worker_page {
  width: 100%;
  height: 100vh;
  display: flex;
  padding: 2rem;

  .tree_item {
    width: 20%;
    border-right: 1px solid #eeeeee;
  }

  .file_item {
    display: flex;
    flex: 1;
    justify-content: center;
    align-items: center;

    .file-reader div:first-of-type {
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 20px;

      span {
        font-size: 18px;
        font-weight: 500;
      }
    }

    .non-centered {
      align-items: flex-start;
    }
  }
}

pre {
  text-align: left;
  background-color: #f8f8f8;
  padding: 10px;
  border: 1px solid #ccc;
}
</style>
