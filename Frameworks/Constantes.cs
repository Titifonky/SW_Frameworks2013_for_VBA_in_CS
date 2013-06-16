using System;
using System.Runtime.InteropServices;

namespace Framework_SW2013
{

    #region "Enumérations"

    //Cet attribut permet de combiner les valeurs d'enumération
    [Flags]
    public enum TypeFichier_e
    {
        cAutre = 1,
        cAssemblage = 2,
        cPiece = 4,
        cDessin = 8
    }

    //Cet attribut permet de combiner les valeurs d'enumération
    [Flags]
    public enum TypeCorps_e
    {
        cAucun = 1,
        cTole = 2,
        cBarre = 4,
        cAutre = 8,
        cTous = cTole | cBarre | cAutre
    }

    //Cet attribut permet de combiner les valeurs d'enumération
    [Flags]
    public enum TypeConfig_e
    {
        cDeBase = 1,
        cDerivee = 2,
        cDepliee = 4,
        cPliee = 8,
        cTous = cDeBase | cDerivee | cDepliee | cPliee
    }

    public enum EtatFonction_e
    {
        cDesactivee = 0,
        cActivee = 1
    }

    public enum Orientation_e
    {
        cPortrait = 0,
        cPaysage = 1
    }

    public enum Format_e
    {
        cA0 = 5,
        cA1 = 6,
        cA2 = 7,
        cA3 = 8,
        cA4 = 9,
        cA5 = 10,
        cUtilisateur = 11
    }

    public enum Extension_e
    {
        cDXF = 1,
        cDWG = 2,
        cPDF = 4
    }

    public enum Grugeage_e
    {
        cRectangulaire = 1,
        cDechirureDecoupe = 2,
        cArrondi = 3,
        cAucun = 4,
        cDechirureProlonge = 5
    }

    #endregion

    internal static class CONSTANTES
    {
        #region "Variables locales"

        internal static String CONFIG_DEPLIEE_PATTERN = "^*SM-FLAT-PATTERN*";
        internal static String CONFIG_DEPLIEE = "SM-FLAT-PATTERN";
        internal static String CONFIG_PLIEE_PATTERN = "^[0-9]";
        internal static String ARTICLE_LISTE_DES_PIECES_SOUDEES = "Article-liste-des-pièces-soudées";
        internal static String EPAISSEUR_DE_TOLE = "Epaisseur de la tôle";
        internal static String NO_DOSSIER = "NoDossier";
        internal static String NO_CONFIG = "NoConfig";
        internal static String NOM_ELEMENT = "Element";
        internal static String PROFIL_NOM = "Profil";
        internal static String PROFIL_ANGLE1 = "ANGLE1";
        internal static String PROFIL_ANGLE2 = "ANGLE2";
        internal static String PROFIL_LONGUEUR = "LONGUEUR";
        internal static String PROFIL_MASSE = "Masse";
        internal static String PROFIL_MATERIAU = "MATERIAL";
        internal static String CUBE_DE_VISUALISATION = "Cube de visualisation";
        internal static String MODELE_DE_DESSIN_LASER = "MacroLaser";
        internal static String NOM_CORPS_DEPLIEE = "Etat déplié";
        internal static String ETAT_D_AFFICHAGE = "Etat d'affichage-";
        #endregion
    }

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("EE3D32C7-E632-4DCE-83F7-CF58F569F523")]
    public interface IConstantes
    {
        String ARTICLE_LISTE_DES_PIECES_SOUDEES { get; }
        String CONFIG_DEPLIEE { get; }
        String EPAISSEUR_DE_TOLE { get; }
        String NO_DOSSIER { get; }
        String NO_CONFIG { get; }
        String NOM_ELEMENT { get; }
        String MODELE_DE_DESSIN_LASER { get; }
        String NOM_CORPS_DEPLIEE { get; }
        String ETAT_D_AFFICHAGE { get; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("5015C3D8-9DC1-4B38-BA40-E04AEA31A45A")]
    [ProgId("Frameworks.Constantes")]
    public class Constantes : IConstantes
    {
        #region "Propriétés"

        public String ARTICLE_LISTE_DES_PIECES_SOUDEES { get { return CONSTANTES.ARTICLE_LISTE_DES_PIECES_SOUDEES; } }
        public String CONFIG_DEPLIEE { get { return CONSTANTES.CONFIG_DEPLIEE; } }
        public String EPAISSEUR_DE_TOLE { get { return CONSTANTES.EPAISSEUR_DE_TOLE; } }
        public String NO_DOSSIER { get { return CONSTANTES.NO_DOSSIER; } }
        public String NO_CONFIG { get { return CONSTANTES.NO_CONFIG; } }
        public String NOM_ELEMENT { get { return CONSTANTES.NOM_ELEMENT; } }
        public String MODELE_DE_DESSIN_LASER { get { return CONSTANTES.MODELE_DE_DESSIN_LASER; } }
        public String NOM_CORPS_DEPLIEE { get { return CONSTANTES.NOM_CORPS_DEPLIEE; } }
        public String ETAT_D_AFFICHAGE { get { return CONSTANTES.ETAT_D_AFFICHAGE; } }

        #endregion
    }

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("E2C9A545-2DBF-4F93-BB30-5EC7D41A5D4E")]
    public interface IswFeatureType_e
    {
        String swTnExplodeLineProfileFeature { get; }
        String swTnInContextFeatHolder { get; }
        String swTnMateCoincident { get; }
        String swTnMateConcentric { get; }
        String swTnMateDistanceDim { get; }
        String swTnMateGroup { get; }
        String swTnMateInPlace { get; }
        String swTnMateParallel { get; }
        String swTnMatePerpendicular { get; }
        String swTnMatePlanarAngleDim { get; }
        String swTnMateSymmetric { get; }
        String swTnMateTangent { get; }
        String swTnMateWidth { get; }
        String swTnReference { get; }
        String swTnSmartComponentFeature { get; }
        String swTnBaseBody { get; }
        String swTnBlend { get; }
        String swTnBlendCut { get; }
        String swTnBoss { get; }
        String swTnBossThin { get; }
        String swTnCavity { get; }
        String swTnChamfer { get; }
        String swTnCirPattern { get; }
        String swTnCombineBodies { get; }
        String swTnCosmeticThread { get; }
        String swTnCurvePattern { get; }
        String swTnCut { get; }
        String swTnCutThin { get; }
        String swTnDeform { get; }
        String swTnDeleteBody { get; }
        String swTnDelFace { get; }
        String swTnDerivedCirPattern { get; }
        String swTnDerivedLPattern { get; }
        String swTnDome { get; }
        String swTnDraft { get; }
        String swTnEmboss { get; }
        String swTnExtrusion { get; }
        String swTnFillet { get; }
        String swTnHelix { get; }
        String swTnHoleWzd { get; }
        String swTnImported { get; }
        String swTnICE { get; }
        String swTnLocalCirPattern { get; }
        String swTnLocalLPattern { get; }
        String swTnLPattern { get; }
        String swTnMirrorPattern { get; }
        String swTnMirrorSolid { get; }
        String swTnMoveCopyBody { get; }
        String swTnReplaceFace { get; }
        String swTnRevCut { get; }
        String swTnRevolution { get; }
        String swTnRevolutionThin { get; }
        String swTnShape { get; }
        String swTnShell { get; }
        String swTnSplit { get; }
        String swTnStock { get; }
        String swTnSweep { get; }
        String swTnSweepCut { get; }
        String swTnTablePattern { get; }
        String swTnThicken { get; }
        String swTnThickenCut { get; }
        String swTnVarFillet { get; }
        String swTnVolSweep { get; }
        String swTnVolSweepCut { get; }
        String swTnMirrorStock { get; }
        String swTnAbsoluteView { get; }
        String swTnAlignGroup { get; }
        String swTnAuxiliaryView { get; }
        String swTnBomTableAnchor { get; }
        String swTnBomTableFeature { get; }
        String swTnBomTemplate { get; }
        String swTnBreakLine { get; }
        String swTnCenterMark { get; }
        String swTnDetailCircle { get; }
        String swTnDetailView { get; }
        String swTnDrBreakoutSectionLine { get; }
        String swTnDrSectionLine { get; }
        String swTnDrSheet { get; }
        String swTnDrTemplate { get; }
        String swTnDrViewDetached { get; }
        String swTnGeneralTableAnchor { get; }
        String swTnHoleTableAnchor { get; }
        String swTnHoleTableFeat { get; }
        String swTnHoleTableFeature { get; }
        String swTnLiveSection { get; }
        String swTnRelativeView { get; }
        String swTnRevisionTableAnchor { get; }
        String swTnRevisionTableFeature { get; }
        String swTnSection { get; }
        String swTnSectionAssemView { get; }
        String swTnSectionPartView { get; }
        String swTnSectionView { get; }
        String swTnUnfoldedView { get; }
        String swTnBlocksFolder { get; }
        String swTnCommentsFolder { get; }
        String swTnCutListFolder { get; }
        String swTnDocsFolder { get; }
        String swTnFeatSolidBodyFolder { get; }
        String swTnFeatSurfaceBodyFolder { get; }
        String swTnFeatureFolder { get; }
        String swTnFtrFolder { get; }
        String swTnGridDetailFolder { get; }
        String swTnInsertedFeatureFolder { get; }
        String swTnLiveSectionFolder { get; }
        String swTnMateReferenceGroupFolder { get; }
        String swTnMaterialFolder { get; }
        String swTnPosGroupFolder { get; }
        String swTnProfileFtrFolder { get; }
        String swTnRefAxisFtrFolder { get; }
        String swTnRefPlaneFtrFolder { get; }
        String swTnSolidBodyFolder { get; }
        String swTnSmartComponentFolder { get; }
        String swTnSmartComponentRefFolder { get; }
        String swTnSubAtomFolder { get; }
        String swTnSubWeldFolder { get; }
        String swTnSurfaceBodyFolder { get; }
        String swTnTableFolder { get; }
        String swTnAttribute { get; }
        String swTnBlockDef { get; }
        String swTnComments { get; }
        String swTnConfigBuilderFeature { get; }
        String swTnConfiguration { get; }
        String swTnCurveInFile { get; }
        String swTnDesignTableFeature { get; }
        String swTnDetailCabinet { get; }
        String swTnEmbedLinkDoc { get; }
        String swTnGridFeature { get; }
        String swTnJournal { get; }
        String swTnLibraryFeature { get; }
        String swTnPartConfiguration { get; }
        String swTnReferenceBrowser { get; }
        String swTnReferenceEmbedded { get; }
        String swTnReferenceInternal { get; }
        String swTnScale { get; }
        String swTnViewerBodyFeature { get; }
        String swTnXMLRulesFeature { get; }
        String swTnMoldCoreCavitySolids { get; }
        String swTnMoldPartingGeom { get; }
        String swTnMoldPartingLine { get; }
        String swTnMoldShutOffSrf { get; }
        String swTnAEM3DContact { get; }
        String swTnAEMGravity { get; }
        String swTnAEMLinearDamper { get; }
        String swTnAEMLinearForce { get; }
        String swTnAEMLinearMotor { get; }
        String swTnAEMLinearSpring { get; }
        String swTnAEMRotationalMotor { get; }
        String swTnAEMSimFeature { get; }
        String swTnAEMTorque { get; }
        String swTnAEMTorsionalDamper { get; }
        String swTnAEMTorsionalSpring { get; }
        String swTnCoordinateSystem { get; }
        String swTnCoordSys { get; }
        String swTnRefAxis { get; }
        String swTnReferenceCurve { get; }
        String swTnRefPlane { get; }
        String swTnAmbientLight { get; }
        String swTnCameraFeature { get; }
        String swTnDirectionLight { get; }
        String swTnPointLight { get; }
        String swTnSpotLight { get; }
        String swTnBaseFlange { get; }
        String swTnBending { get; }
        String swTnBreakCorner { get; }
        String swTnCornerTrim { get; }
        String swTnCrossBreak { get; }
        String swTnEdgeFlange { get; }
        String swTnFlatPattern { get; }
        String swTnFlattenBends { get; }
        String swTnFold { get; }
        String swTnFormToolInstance { get; }
        String swTnHem { get; }
        String swTnLoftedBend { get; }
        String swTnOneBend { get; }
        String swTnProcessBends { get; }
        String swTnSheetMetal { get; }
        String swTnSketchBend { get; }
        String swTnSM3dBend { get; }
        String swTnSMBaseFlange { get; }
        String swTnSolidToSheetMetal { get; }
        String swTnSMMiteredFlange { get; }
        String swTnToroidalBend { get; }
        String swTnUiBend { get; }
        String swTnUnFold { get; }
        String swTn3DProfileFeature { get; }
        String swTn3DSplineCurve { get; }
        String swTnCompositeCurve { get; }
        String swTnLayoutProfileFeature { get; }
        String swTnPLine { get; }
        String swTnProfileFeature { get; }
        String swTnRefCurve { get; }
        String swTnRefPoint { get; }
        String swTnSketchBlockDefinition { get; }
        String swTnSketchHole { get; }
        String swTnSketchPattern { get; }
        String swTnSketchPicture { get; }
        String swTnBlendRefSurface { get; }
        String swTnFillRefSurface { get; }
        String swTnMidRefSurface { get; }
        String swTnOffsetRefSurface { get; }
        String swTnRadiateRefSurface { get; }
        String swTnRefSurface { get; }
        String swTnRevolvRefSurf { get; }
        String swTnRuledSrfFromEdge { get; }
        String swTnSewRefSurface { get; }
        String swTnSurfCut { get; }
        String swTnSweepRefSurface { get; }
        String swTnTrimRefSurface { get; }
        String swTnUnTrimRefSurf { get; }
        String swTnVolSweepRefSurface { get; }
        String swTnGusset { get; }
        String swTnWeldMemberFeat { get; }
        String swTnWeldmentFeature { get; }
        String swTnWeldmentTableAnchor { get; }
        String swTnWeldmentTableFeature { get; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("9D843392-D00C-4D2A-B8D2-1AC60A99042A")]
    [ProgId("Frameworks.swFeatureType_e")]
    public class swFeatureType_e : IswFeatureType_e
    {
        #region "Propriétés"

        public String swTnExplodeLineProfileFeature { get { return "ExplodeLineProfileFeature"; } }
        public String swTnInContextFeatHolder { get { return "InContextFeatHolder"; } }
        public String swTnMateCoincident { get { return "MateCoincident"; } }
        public String swTnMateConcentric { get { return "MateConcentric"; } }
        public String swTnMateDistanceDim { get { return "MateDistanceDim"; } }
        public String swTnMateGroup { get { return "MateGroup"; } }
        public String swTnMateInPlace { get { return "MateInPlace"; } }
        public String swTnMateParallel { get { return "MateParallel"; } }
        public String swTnMatePerpendicular { get { return "MatePerpendicular"; } }
        public String swTnMatePlanarAngleDim { get { return "MatePlanarAngleDim"; } }
        public String swTnMateSymmetric { get { return "MateSymmetric"; } }
        public String swTnMateTangent { get { return "MateTangent"; } }
        public String swTnMateWidth { get { return "MateWidth"; } }
        public String swTnReference { get { return "Reference"; } }
        public String swTnSmartComponentFeature { get { return "SmartComponentFeature"; } }
        public String swTnBaseBody { get { return "BaseBody"; } }
        public String swTnBlend { get { return "Blend"; } }
        public String swTnBlendCut { get { return "BlendCut"; } }
        public String swTnBoss { get { return "Boss"; } }
        public String swTnBossThin { get { return "BossThin"; } }
        public String swTnCavity { get { return "Cavity"; } }
        public String swTnChamfer { get { return "Chamfer"; } }
        public String swTnCirPattern { get { return "CirPattern"; } }
        public String swTnCombineBodies { get { return "CombineBodies"; } }
        public String swTnCosmeticThread { get { return "CosmeticThread"; } }
        public String swTnCurvePattern { get { return "CurvePattern"; } }
        public String swTnCut { get { return "Cut"; } }
        public String swTnCutThin { get { return "CutThin"; } }
        public String swTnDeform { get { return "Deform"; } }
        public String swTnDeleteBody { get { return "DeleteBody"; } }
        public String swTnDelFace { get { return "DelFace"; } }
        public String swTnDerivedCirPattern { get { return "DerivedCirPattern"; } }
        public String swTnDerivedLPattern { get { return "DerivedLPattern"; } }
        public String swTnDome { get { return "Dome"; } }
        public String swTnDraft { get { return "Draft"; } }
        public String swTnEmboss { get { return "Emboss"; } }
        public String swTnExtrusion { get { return "Extrusion"; } }
        public String swTnFillet { get { return "Fillet"; } }
        public String swTnHelix { get { return "Helix"; } }
        public String swTnHoleWzd { get { return "HoleWzd"; } }
        public String swTnImported { get { return "Imported"; } }
        public String swTnICE { get { return "ICE"; } }
        public String swTnLocalCirPattern { get { return "LocalCirPattern"; } }
        public String swTnLocalLPattern { get { return "LocalLPattern"; } }
        public String swTnLPattern { get { return "LPattern"; } }
        public String swTnMirrorPattern { get { return "MirrorPattern"; } }
        public String swTnMirrorSolid { get { return "MirrorSolid"; } }
        public String swTnMoveCopyBody { get { return "MoveCopyBody"; } }
        public String swTnReplaceFace { get { return "ReplaceFace"; } }
        public String swTnRevCut { get { return "RevCut"; } }
        public String swTnRevolution { get { return "Revolution"; } }
        public String swTnRevolutionThin { get { return "RevolutionThin"; } }
        public String swTnShape { get { return "Shape"; } }
        public String swTnShell { get { return "Shell"; } }
        public String swTnSplit { get { return "Split"; } }
        public String swTnStock { get { return "Stock"; } }
        public String swTnSweep { get { return "Sweep"; } }
        public String swTnSweepCut { get { return "SweepCut"; } }
        public String swTnTablePattern { get { return "TablePattern"; } }
        public String swTnThicken { get { return "Thicken"; } }
        public String swTnThickenCut { get { return "ThickenCut"; } }
        public String swTnVarFillet { get { return "VarFillet"; } }
        public String swTnVolSweep { get { return "VolSweep"; } }
        public String swTnVolSweepCut { get { return "VolSweepCut"; } }
        public String swTnMirrorStock { get { return "MirrorStock"; } }
        public String swTnAbsoluteView { get { return "AbsoluteView"; } }
        public String swTnAlignGroup { get { return "AlignGroup"; } }
        public String swTnAuxiliaryView { get { return "AuxiliaryView"; } }
        public String swTnBomTableAnchor { get { return "BomTemplate"; } }
        public String swTnBomTableFeature { get { return "BomFeat"; } }
        public String swTnBomTemplate { get { return "BomTemplate"; } }
        public String swTnBreakLine { get { return "BreakLine"; } }
        public String swTnCenterMark { get { return "CenterMark"; } }
        public String swTnDetailCircle { get { return "DetailCircle"; } }
        public String swTnDetailView { get { return "DetailView"; } }
        public String swTnDrBreakoutSectionLine { get { return "DrBreakoutSectionLine"; } }
        public String swTnDrSectionLine { get { return "DrSectionLine"; } }
        public String swTnDrSheet { get { return "DrSheet"; } }
        public String swTnDrTemplate { get { return "DrTemplate"; } }
        public String swTnDrViewDetached { get { return "DrViewDetached"; } }
        public String swTnGeneralTableAnchor { get { return "GeneralTableAnchor"; } }
        public String swTnHoleTableAnchor { get { return "HoleTableAnchor"; } }
        public String swTnHoleTableFeat { get { return "HoleTableFeat"; } }
        public String swTnHoleTableFeature { get { return "HoleTableFeat"; } }
        public String swTnLiveSection { get { return "LiveSection"; } }
        public String swTnRelativeView { get { return "RelativeView"; } }
        public String swTnRevisionTableAnchor { get { return "RevisionTableAnchor"; } }
        public String swTnRevisionTableFeature { get { return "RevisionTableFeat"; } }
        public String swTnSection { get { return "Section"; } }
        public String swTnSectionAssemView { get { return "SectionAssemView"; } }
        public String swTnSectionPartView { get { return "SectionPartView"; } }
        public String swTnSectionView { get { return "SectionView"; } }
        public String swTnUnfoldedView { get { return "UnfoldedView"; } }
        public String swTnBlocksFolder { get { return "BlockFolder"; } }
        public String swTnCommentsFolder { get { return "CommentsFolder"; } }
        public String swTnCutListFolder { get { return "CutListFolder"; } }
        public String swTnDocsFolder { get { return "DocsFolder"; } }
        public String swTnFeatSolidBodyFolder { get { return "FeatSolidBodyFolder"; } }
        public String swTnFeatSurfaceBodyFolder { get { return "FeatSurfaceBodyFolder"; } }
        public String swTnFeatureFolder { get { return "FtrFolder"; } }
        public String swTnFtrFolder { get { return "FtrFolder"; } }
        public String swTnGridDetailFolder { get { return "GridDetailFolder"; } }
        public String swTnInsertedFeatureFolder { get { return "InsertedFeatureFolder"; } }
        public String swTnLiveSectionFolder { get { return "LiveSectionFolder"; } }
        public String swTnMateReferenceGroupFolder { get { return "MateReferenceGroupFolder"; } }
        public String swTnMaterialFolder { get { return "MaterialFolder"; } }
        public String swTnPosGroupFolder { get { return "PosGroupFolder"; } }
        public String swTnProfileFtrFolder { get { return "ProfileFtrFolder"; } }
        public String swTnRefAxisFtrFolder { get { return "RefAxisFtrFolder"; } }
        public String swTnRefPlaneFtrFolder { get { return "RefPlaneFtrFolder"; } }
        public String swTnSolidBodyFolder { get { return "SolidBodyFolder"; } }
        public String swTnSmartComponentFolder { get { return "SmartComponentFolder"; } }
        public String swTnSmartComponentRefFolder { get { return "SmartComponentRefFolder"; } }
        public String swTnSubAtomFolder { get { return "SubAtomFolder"; } }
        public String swTnSubWeldFolder { get { return "SubWeldFolder"; } }
        public String swTnSurfaceBodyFolder { get { return "SurfaceBodyFolder"; } }
        public String swTnTableFolder { get { return "TableFolder"; } }
        public String swTnAttribute { get { return "Attribute"; } }
        public String swTnBlockDef { get { return "BlockDef"; } }
        public String swTnComments { get { return "Comments"; } }
        public String swTnConfigBuilderFeature { get { return "ConfigBuilderFeature"; } }
        public String swTnConfiguration { get { return "Configuration"; } }
        public String swTnCurveInFile { get { return "CurveInFile"; } }
        public String swTnDesignTableFeature { get { return "DesignTableFeature"; } }
        public String swTnDetailCabinet { get { return "DetailCabinet"; } }
        public String swTnEmbedLinkDoc { get { return "EmbedLinkDoc"; } }
        public String swTnGridFeature { get { return "GridFeature"; } }
        public String swTnJournal { get { return "Journal"; } }
        public String swTnLibraryFeature { get { return "LibraryFeature"; } }
        public String swTnPartConfiguration { get { return "PartConfiguration"; } }
        public String swTnReferenceBrowser { get { return "ReferenceBrowser"; } }
        public String swTnReferenceEmbedded { get { return "ReferenceEmbedded"; } }
        public String swTnReferenceInternal { get { return "ReferenceInternal"; } }
        public String swTnScale { get { return "Scale"; } }
        public String swTnViewerBodyFeature { get { return "ViewerBodyFeature"; } }
        public String swTnXMLRulesFeature { get { return "XMLRulesFeature"; } }
        public String swTnMoldCoreCavitySolids { get { return "MoldCoreCavitySolids"; } }
        public String swTnMoldPartingGeom { get { return "MoldPartingGeom"; } }
        public String swTnMoldPartingLine { get { return "MoldPartingLine"; } }
        public String swTnMoldShutOffSrf { get { return "MoldShutOffSrf"; } }
        public String swTnAEM3DContact { get { return "AEM3DContact"; } }
        public String swTnAEMGravity { get { return "AEMGravity"; } }
        public String swTnAEMLinearDamper { get { return "AEMLinearDamper"; } }
        public String swTnAEMLinearForce { get { return "AEMLinearForce"; } }
        public String swTnAEMLinearMotor { get { return "AEMLinearMotor"; } }
        public String swTnAEMLinearSpring { get { return "AEMLinearSpring"; } }
        public String swTnAEMRotationalMotor { get { return "AEMRotationalMotor"; } }
        public String swTnAEMSimFeature { get { return "AEMSimFeature"; } }
        public String swTnAEMTorque { get { return "AEMTorque"; } }
        public String swTnAEMTorsionalDamper { get { return "AEMTorsionalDamper"; } }
        public String swTnAEMTorsionalSpring { get { return "AEMTorsionalSpring"; } }
        public String swTnCoordinateSystem { get { return "CoordSys"; } }
        public String swTnCoordSys { get { return "CoordSys"; } }
        public String swTnRefAxis { get { return "RefAxis"; } }
        public String swTnReferenceCurve { get { return "ReferenceCurve"; } }
        public String swTnRefPlane { get { return "RefPlane"; } }
        public String swTnAmbientLight { get { return "AmbientLight"; } }
        public String swTnCameraFeature { get { return "CameraFeature"; } }
        public String swTnDirectionLight { get { return "DirectionLight"; } }
        public String swTnPointLight { get { return "PointLight"; } }
        public String swTnSpotLight { get { return "SpotLight"; } }
        public String swTnBaseFlange { get { return "SMBaseFlange"; } }
        public String swTnSolidToSheetMetal { get { return "SolidToSheetMetal"; } }
        public String swTnBending { get { return "Bending"; } }
        public String swTnBreakCorner { get { return "BreakCorner"; } }
        public String swTnCornerTrim { get { return "CornerTrim"; } }
        public String swTnCrossBreak { get { return "CrossBreak"; } }
        public String swTnEdgeFlange { get { return "EdgeFlange"; } }
        public String swTnFlatPattern { get { return "FlatPattern"; } }
        public String swTnFlattenBends { get { return "FlattenBends"; } }
        public String swTnFold { get { return "Fold"; } }
        public String swTnFormToolInstance { get { return "FormToolInstance"; } }
        public String swTnHem { get { return "Hem"; } }
        public String swTnLoftedBend { get { return "LoftedBend"; } }
        public String swTnOneBend { get { return "OneBend"; } }
        public String swTnProcessBends { get { return "ProcessBends"; } }
        public String swTnSheetMetal { get { return "SheetMetal"; } }
        public String swTnSketchBend { get { return "SketchBend"; } }
        public String swTnSM3dBend { get { return "SM3dBend"; } }
        public String swTnSMBaseFlange { get { return "SMBaseFlange"; } }
        public String swTnSMMiteredFlange { get { return "SMMiteredFlange"; } }
        public String swTnToroidalBend { get { return "ToroidalBend"; } }
        public String swTnUiBend { get { return "UiBend"; } }
        public String swTnUnFold { get { return "UnFold"; } }
        public String swTn3DProfileFeature { get { return "3DProfileFeature"; } }
        public String swTn3DSplineCurve { get { return "3DSplineCurve"; } }
        public String swTnCompositeCurve { get { return "CompositeCurve"; } }
        public String swTnLayoutProfileFeature { get { return "LayoutProfileFeature"; } }
        public String swTnPLine { get { return "PLine"; } }
        public String swTnProfileFeature { get { return "ProfileFeature"; } }
        public String swTnRefCurve { get { return "RefCurve"; } }
        public String swTnRefPoint { get { return "RefPoint"; } }
        public String swTnSketchBlockDefinition { get { return "SketchBlockDef"; } }
        public String swTnSketchHole { get { return "SketchHole"; } }
        public String swTnSketchPattern { get { return "SketchPattern"; } }
        public String swTnSketchPicture { get { return "SketchBitmap"; } }
        public String swTnBlendRefSurface { get { return "BlendRefSurface"; } }
        public String swTnFillRefSurface { get { return "FillRefSurface"; } }
        public String swTnMidRefSurface { get { return "MidRefSurface"; } }
        public String swTnOffsetRefSurface { get { return "OffsetRefSurface"; } }
        public String swTnRadiateRefSurface { get { return "RadiateRefSurface"; } }
        public String swTnRefSurface { get { return "RefSurface"; } }
        public String swTnRevolvRefSurf { get { return "RevolvRefSurf"; } }
        public String swTnRuledSrfFromEdge { get { return "RuledSrfFromEdge"; } }
        public String swTnSewRefSurface { get { return "SewRefSurface"; } }
        public String swTnSurfCut { get { return "SurfCut"; } }
        public String swTnSweepRefSurface { get { return "SweepRefSurface"; } }
        public String swTnTrimRefSurface { get { return "TrimRefSurface"; } }
        public String swTnUnTrimRefSurf { get { return "UnTrimRefSurf"; } }
        public String swTnVolSweepRefSurface { get { return "VolSweepRefSurface"; } }
        public String swTnGusset { get { return "Gusset"; } }
        public String swTnWeldMemberFeat { get { return "WeldMemberFeat"; } }
        public String swTnWeldmentFeature { get { return "WeldmentFeature"; } }
        public String swTnWeldmentTableAnchor { get { return "WeldmentTableAnchor"; } }
        public String swTnWeldmentTableFeature { get { return "WeldmentTableFeat"; } }

        #endregion
    }
}
