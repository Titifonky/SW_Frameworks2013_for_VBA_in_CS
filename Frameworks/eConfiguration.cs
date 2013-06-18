﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("93F8AAEE-5820-11E2-A4D1-91046188709B")]
    public interface IeConfiguration
    {
        Configuration SwConfiguration { get; }
        eModele Modele { get; }
        String Nom { get; set; }
        TypeConfig_e TypeConfig { get; }
        eConfiguration ConfigurationParent { get; }
        eConfiguration ConfigurationRacine { get; }
        eGestDeProprietes GestDeProprietes { get; }
        Boolean SupprimerLesNouvellesFonctions { get; set; }
        Boolean SupprimerLesNouveauxComposants { get; set; }
        Boolean CacherLesNouveauxComposant { get; set; }
        Boolean AfficherLesNouveauxComposantDansLaNomenclature { get; set; }
        Boolean Est(TypeConfig_e T);
        Boolean Activer();
        Boolean Supprimer();
        ArrayList ConfigurationsEnfants(String NomConfiguration = "", TypeConfig_e TypeDeLaConfig = TypeConfig_e.cTous);
        eConfiguration AjouterUneConfigurationDerivee(String NomConfigDerivee);
        void RenommerEtatAffichage();
        eCorps CorpsDeplie();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("9818CA82-5820-11E2-852A-92046188709B")]
    [ProgId("Frameworks.eConfiguration")]
    public class eConfiguration : IeConfiguration, IComparable<eConfiguration>, IComparer<eConfiguration>, IEquatable<eConfiguration>
    {
#region "Variables locales"

        private Boolean _EstInitialise = false;

        private Configuration _SwConfiguration = null;
        private eModele _Modele = null;
        private int _ID = 0;

#endregion

#region "Constructeur\Destructeur"

        public eConfiguration() { }

#endregion

#region "Propriétés"

        /// <summary>
        /// Retourne l'objet Configuration associé.
        /// </summary>
        public Configuration SwConfiguration
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                
                if (_SwConfiguration == null)
                {
                    foreach (String pNomConfig in _Modele.SwModele.GetConfigurationNames())
                    {
                        Configuration pSwConfig = _Modele.SwModele.GetConfigurationByName(pNomConfig);
                        if (pSwConfig.GetID() == _ID)
                        {
                            _SwConfiguration = pSwConfig;
                            break;
                        }
                    }
                }

                return _SwConfiguration;
            }
        }

        /// <summary>
        /// Retourne le parent ExtModele.
        /// </summary>
        public eModele Modele { get { Debug.Info(MethodBase.GetCurrentMethod()); return _Modele; } }

        /// <summary>
        /// Retourne ou défini le nom de la configuration.
        /// </summary>
        public String Nom { get { Debug.Info(MethodBase.GetCurrentMethod()); return SwConfiguration.Name; } set { Debug.Info(MethodBase.GetCurrentMethod()); SwConfiguration.Name = value; } }

        /// <summary>
        /// Retourne le type de la configuration.
        /// </summary>
        public TypeConfig_e TypeConfig
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                TypeConfig_e T = 0;
                if (Regex.IsMatch(SwConfiguration.Name, CONSTANTES.CONFIG_DEPLIEE_PATTERN))
                    T = TypeConfig_e.cDepliee;
                else if (Regex.IsMatch(SwConfiguration.Name, CONSTANTES.CONFIG_PLIEE_PATTERN))
                {
                    T = TypeConfig_e.cPliee;
                }

                if (SwConfiguration.IsDerived() != false)
                    T |= TypeConfig_e.cDerivee;
                else
                    T |= TypeConfig_e.cRacine;

                if (!T.HasFlag(TypeConfig_e.cDepliee))
                    T |= TypeConfig_e.cDeBase;

                return T;
            }
        }

        /// <summary>
        /// Retourne la configuration parent.
        /// </summary>
        public eConfiguration ConfigurationParent
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                eConfiguration pConfigParent = new eConfiguration();
                Configuration pSwConfigurationParent = null;
                if (SwConfiguration.IsDerived() == true)
                    pSwConfigurationParent = SwConfiguration.GetParent();

                if ((pSwConfigurationParent != null) && pConfigParent.Init(pSwConfigurationParent, _Modele))
                    return pConfigParent;

                return null;
            }
        }

        /// <summary>
        /// Retourne la configuration racine, c'est la configuration parent la plus haute.
        /// </summary>
        public eConfiguration ConfigurationRacine
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                // Si elle est derivée, on lance la recherche
                if (Est(TypeConfig_e.cDerivee))
                {
                    eConfiguration pConfig = ConfigurationParent;

                    // Tant que pConfig est derivee, on boucle.
                    while (pConfig.Est(TypeConfig_e.cDerivee))
                    {
                        pConfig = pConfig.ConfigurationParent;
                    }

                    // Et on renvoi la dernière
                    return pConfig;
                }

                // Si elle n'est pas derivée, on la renvoi
                return this;
            }
        }

        /// <summary>
        /// Retourne le getionnaire de propriétés.
        /// </summary>
        public eGestDeProprietes GestDeProprietes
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                eGestDeProprietes pGestProps = new eGestDeProprietes();
                if (pGestProps.Init(SwConfiguration.CustomPropertyManager, _Modele))
                    return pGestProps;

                return null;
            }
        }


        /// <summary>
        /// Supprimer les nouvelles fonctions.
        /// </summary>
        public Boolean SupprimerLesNouvellesFonctions
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return SwConfiguration.SuppressNewFeatures;
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                SwConfiguration.SuppressNewFeatures = value;
            }
        }

        /// <summary>
        /// Supprimer les nouveaux composants
        /// </summary>
        public Boolean SupprimerLesNouveauxComposants
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return SwConfiguration.SuppressNewComponentModels;
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                SwConfiguration.SuppressNewComponentModels = value;
            }
        }

        /// <summary>
        /// Cacher les nouveaux composants.
        /// </summary>
        public Boolean CacherLesNouveauxComposant
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return SwConfiguration.HideNewComponentModels;
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                SwConfiguration.HideNewComponentModels = value;
            }
        }

        /// <summary>
        /// Afficher les nouveaux composants dans la nomenclature
        /// </summary>
        public Boolean AfficherLesNouveauxComposantDansLaNomenclature
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return SwConfiguration.ShowChildComponentsInBOM;
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                SwConfiguration.ShowChildComponentsInBOM = value;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtConfiguration.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

#endregion

#region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtConfiguration.
        /// </summary>
        /// <param name="Config"></param>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(Configuration Config, eModele Modele)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((Config != null) && (Modele != null) && Modele.EstInitialise)
            {
                _SwConfiguration = Config;
                _Modele = Modele;
                _ID = _SwConfiguration.GetID();

                Debug.Info(this.Nom);
                _EstInitialise = true;
            }
            else
            {
                Debug.Info("!!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        /// <summary>
        /// Test si la configuration est de type T.
        /// </summary>
        /// <param name="T"></param>
        /// <returns></returns>
        public Boolean Est(TypeConfig_e T)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            foreach (TypeConfig_e Tst in Enum.GetValues(typeof(TypeConfig_e)))
            {
                if (T.HasFlag(Tst) && TypeConfig.HasFlag(Tst))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Active la configuration.
        /// </summary>
        /// <returns></returns>
        public Boolean Activer()
        {
            Debug.Info(MethodBase.GetCurrentMethod());
            return Convert.ToBoolean(_Modele.SwModele.ShowConfiguration2(Nom));
        }

        /// <summary>
        /// Supprime la configuration.
        /// </summary>
        /// <returns></returns>
        public Boolean Supprimer()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            Boolean UneConfig = (Modele.SwModele.GetConfigurationCount() == 1);
            String[] pTabNomAff = _SwConfiguration.GetDisplayStates();

            if (pTabNomAff.Length  > 0)
            {
                String pAffichageCourant = pTabNomAff[0];
                foreach (String pNomEtatAffichage in pTabNomAff)
                {
                    if (!(UneConfig && (pNomEtatAffichage == pAffichageCourant)))
                        _SwConfiguration.DeleteDisplayState(pNomEtatAffichage);
                }
            }

            if (!UneConfig)
            {
                _EstInitialise = false;
                return _Modele.SwModele.DeleteConfiguration2(Nom);
            }

            return false;
        }

        /// <summary>
        /// Méthode interne.
        /// Renvoi la liste des configurations enfants filtrée par les arguments.
        /// </summary>
        /// <param name="NomConfiguration"></param>
        /// <param name="TypeDeLaConfig"></param>
        /// <returns></returns>
        internal List<eConfiguration> ListConfigurationsEnfants(String NomConfiguration = "", TypeConfig_e TypeDeLaConfig = TypeConfig_e.cTous)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eConfiguration> pListe = new List<eConfiguration>();

            if (SwConfiguration.GetChildrenCount() == 0)
                return pListe;

            foreach (Configuration pSwConfiguration in SwConfiguration.GetChildren())
            {
                if (Regex.IsMatch(pSwConfiguration.Name, NomConfiguration))
                {
                    eConfiguration pConfig = new eConfiguration();
                    pConfig.Init(pSwConfiguration, Modele);

                    if (pConfig.EstInitialise && pConfig.Est(TypeDeLaConfig))
                        pListe.Add(pConfig);
                }
            }

            return pListe;

        }

        /// <summary>
        /// Renvoi la liste des configurations enfants filtrée par les arguments.
        /// </summary>
        /// <param name="NomConfiguration"></param>
        /// <param name="TypeDeLaConfig"></param>
        /// <returns></returns>
        public ArrayList ConfigurationsEnfants(String NomConfiguration = "", TypeConfig_e TypeDeLaConfig = TypeConfig_e.cTous)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eConfiguration> pListeConfigs = ListConfigurationsEnfants(NomConfiguration, TypeDeLaConfig);
            ArrayList pArrayConfigs = new ArrayList();

            if (pListeConfigs.Count > 0)
                pArrayConfigs = new ArrayList(pListeConfigs);

            return pArrayConfigs;
        }

        /// <summary>
        /// Ajoute une configuration dérivée.
        /// </summary>
        /// <param name="NomConfigDerivee"></param>
        /// <returns></returns>
        public eConfiguration AjouterUneConfigurationDerivee(String NomConfigDerivee)
        {
            Debug.Info(MethodBase.GetCurrentMethod());
            eConfiguration pConfig = new eConfiguration();
            Configuration pSwConfig = _Modele.SwModele.ConfigurationManager.AddConfiguration(NomConfigDerivee, NomConfigDerivee, "", 0, Nom, "");

            if (pConfig.Init(pSwConfig, _Modele))
                return pConfig;

            return null;
        }

        /// <summary>
        /// Renomme les états d'affichages associés à la configuration
        /// </summary>
        public void RenommerEtatAffichage()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            int Index = 1;

            if (_SwConfiguration.GetDisplayStatesCount() > 0)
            {
                foreach (String pNomEtatAffichage in _SwConfiguration.GetDisplayStates())
                {
                    if (Regex.IsMatch(pNomEtatAffichage, "^" + CONSTANTES.ETAT_D_AFFICHAGE))
                    {
                        _SwConfiguration.RenameDisplayState(pNomEtatAffichage, Nom + "_" + Index.ToString());
                        Index++;
                    }
                }
            }
        }

        /// <summary>
        /// Renvoi le corps déplié d'une configuration dépliée
        /// </summary>
        /// <returns></returns>
        public eCorps CorpsDeplie()
        {
            // Si c'est une configuration dépliée
            if (Est(TypeConfig_e.cDepliee))
            {
                Activer();
                // On recherche le corps déplié
                List<eCorps> pListeCorps = Modele.Piece.ListListeDesCorps(CONSTANTES.NOM_CORPS_DEPLIEE);
                return pListeCorps[0];
            }

            return null;
        }

#endregion

#region "Interfaces génériques"

        public int CompareTo(eConfiguration Conf)
        {
            String Nom1 = _Modele.SwModele.GetPathName() + SwConfiguration.Name;
            String Nom2 = Conf._Modele.SwModele.GetPathName() + Conf.SwConfiguration.Name;
            return Nom1.CompareTo(Nom2);
        }

        public int Compare(eConfiguration Conf1, eConfiguration Conf2)
        {
            String Nom1 = Conf1._Modele.SwModele.GetPathName() + Conf1.SwConfiguration.Name;
            String Nom2 = Conf2._Modele.SwModele.GetPathName() + Conf2.SwConfiguration.Name;
            return Nom1.CompareTo(Nom2);
        }

        public Boolean Equals(eConfiguration Conf)
        {
            String Nom1 = _Modele.SwModele.GetPathName() + SwConfiguration.Name;
            String Nom2 = Conf._Modele.SwModele.GetPathName() + Conf.SwConfiguration.Name;
            return Nom1.Equals(Nom2);
        }

#endregion
    }
}
