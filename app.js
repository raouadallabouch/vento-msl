/**
 * VENTO MSL - Core Application Logic
 * Architecture: Zero-dependency, offline-ready, cardless minimalist UI
 * Features: Bilingual (FR/AR) with RTL, interactive feature explanations,
 *           auto-vendor isolation, peer recommendation pitch & multi-sheet Excel export.
 */

(function () {
  'use strict';

  // =========================================================================
  // BILINGUAL TRANSLATIONS & FEATURE EXPLANATIONS
  // =========================================================================
  const i18n = {
    fr: {
      appSubtitle: 'Pilotage Commercial & Couverture Terrain',
      btnInstallApp: "Installer l'App",
      btnImport: 'Importer Excel',
      btnExport: 'Exporter Excel',
      dropTitle: 'Glissez votre fichier de Suivi Excel ici',
      dropSubtitle: 'Formats acceptés : .xlsx, .xls (Fichier volumineux multi-vendeurs ou individuel)',
      btnSelectFile: 'Sélectionner le fichier Excel (.xlsx)',
      vendorLabel: 'Vendeur :',
      pinProfile: 'Mémoriser mon profil',
      pinnedProfile: 'Profil par défaut (Mémorisé)',
      tabClients: 'Clients & Ventes',
      tabCatalog: 'Catalogue SKUs',
      tabLeaderboard: 'Classement Agence',
      kpiRevenue: "Chiffre d'Affaires",
      kpiAvgBasket: 'Panier moyen :',
      kpiMslRate: "Taux d'Activation MSL",
      kpiClients: 'Portefeuille Clients',
      kpiActive: 'actifs',
      kpiDormant: 'dormants (0 DH)',
      kpiOpportunities: 'Opportunités Remise (2%)',
      kpiNear100: '≤ 3 pour 2% Remise',
      kpiReached100: 'Remise 2% Validée',
      filterSegment: 'Segment :',
      all: 'Tous',
      filterSort: 'Trier :',
      sortCADesc: 'CA TTC (Plus élevé)',
      sortCAAsc: 'CA TTC (Moins élevé)',
      sortActDesc: 'Taux Activation % (Décroissant)',
      sortActAsc: 'Taux Activation % (Croissant)',
      sortRemAsc: 'SKUs restants (Proches 100%)',
      sortNomAsc: 'Nom Client (A - Z)',
      smartFilterLabel: 'Actions rapides :',
      chipQuickWins: '🚀 Opportunités Rapides (≤3)',
      chipZero: '⚠️ Non Activés (0%)',
      chipMid: 'En Cours (1% - 50%)',
      chipHigh: 'Bien Engagés (>50%)',
      chipFull: '⭐ 100% Validé',
      chipTopCA: '💎 Top 20% CA',
      showing: 'Affichage de',
      clientsOf: 'client(s) sur',
      btnPrint: 'Imprimer / PDF',
      btnCopyWhatsApp: 'Résumé WhatsApp',
      thCode: 'Code',
      thClient: 'Boutique / Client',
      thSegment: 'Segment',
      thCA: 'CA TTC',
      thActivation: 'Activation MSL',
      thRemaining: 'SKUs Restants',
      thDiagnostic: 'Diagnostic',
      noResultsTitle: 'Aucun client trouvé',
      noResultsDesc: 'Essayez de modifier votre recherche ou de réinitialiser les filtres.',
      btnReset: 'Réinitialiser les filtres',
      btnPrev: 'Précédent',
      btnNext: 'Suivant',
      skuMatrixTitle: 'Matrice de Couverture des Références (MSL)',
      skuMatrixDesc: 'Analyse de la pénétration de chaque article dans le portefeuille du vendeur sélectionné.',
      catHuiles: 'Huiles',
      catEntretien: 'Entretien',
      catSoins: 'Soins & Hygiène',
      catAlim: 'Alimentaire',
      thProduct: 'Produit / Référence',
      thCategory: 'Famille',
      thBuyers: 'Magasins Acheteurs',
      thPenetration: 'Taux de Pénétration',
      thTotalSales: 'Volume / Tonnage (Kg)',
      thStatus: 'Statut',
      leaderboardTitle: "Classement des Vendeurs de l'Agence",
      leaderboardDesc: "Comparatif de performance globale du chiffre d'affaires et de taux MSL.",
      modalVendor: 'Vendeur :',
      modalCA: 'CA TTC :',
      modalActivation: "Quota d'Activation MSL",
      modalRemaining: 'SKUs Manquants pour la Remise 2%',
      modalPillarsTitle: 'Éligibilité Remise MSL (2% pour le magasin) : Les 7 Piliers & Règle MTO',
      modalPitchTitle: 'Diagnostic Terrain & Argumentaire de Vente',
      modalPitchTag: 'Priorités Visite',
      modalOrderedSkus: 'Articles Commandés au Catalogue (Tonnage Kg)',
      modalMissingSkus: 'Articles Non Détenus au Catalogue',
      modalSearchOrderedPlaceholder: 'Rechercher article commandé...',
      modalSearchMissingPlaceholder: 'Rechercher article non détenu...',
      modalNoSkuFound: 'Aucun article trouvé',
      modalBtnWhatsApp: 'Partager via WhatsApp',
      modalBtnClose: 'Fermer',
      searchPlaceholder: 'Rechercher par nom de boutique ou code client...',
      skuSearchPlaceholder: 'Filtrer les articles...',
    },
    ar: {
      appSubtitle: 'التتبع التجاري وتغطية نقاط البيع',
      btnInstallApp: 'تثبيت التطبيق',
      btnImport: 'تحميل ملف إكسيل',
      btnExport: 'تصدير إكسيل',
      dropTitle: 'ضع ملف الإكسيل هنا مباشرة',
      dropSubtitle: 'الصيغ المقبولة: .xlsx أو .xls (ملف شامل لجميع البائعين أو فردي)',
      btnSelectFile: 'اختيار ملف إكسيل (.xlsx)',
      vendorLabel: 'البائع :',
      pinProfile: 'حفظ كحسابي الافتراضي',
      pinnedProfile: 'حسابي الافتراضي (محفوظ)',
      tabClients: 'الزبناء والمبيعات',
      tabCatalog: 'دليل المنتجات',
      tabLeaderboard: 'ترتيب الوكالة',
      kpiRevenue: 'المبيعات الكلية (CA)',
      kpiAvgBasket: 'متوسط السلة :',
      kpiMslRate: 'نسبة التفعيل الإلزامي (MSL)',
      kpiClients: 'محفظة المحلات',
      kpiActive: 'محل نشط',
      kpiDormant: 'راكد (0 درهم)',
      kpiOpportunities: 'فرص تأهيل الخصم (2%)',
      kpiNear100: '≤ 3 لخصم 2%',
      kpiReached100: 'مؤهل لخصم 2%',
      filterSegment: 'الصنف :',
      all: 'الكل',
      filterSort: 'ترتيب :',
      sortCADesc: 'المبيعات (الأعلى أولاً)',
      sortCAAsc: 'المبيعات (الأقل أولاً)',
      sortActDesc: 'نسبة التفعيل % (تنازلي)',
      sortActAsc: 'نسبة التفعيل % (تصاعدي)',
      sortRemAsc: 'المتبقي (الأقرب للهدف)',
      sortNomAsc: 'اسم المحل (أ - ي)',
      smartFilterLabel: 'تصفية سريعة :',
      chipQuickWins: '🚀 فرص سريعة (≤3 منتجات)',
      chipZero: '⚠️ غير مفعّلة (0%)',
      chipMid: 'قيد التفعيل (1% - 50%)',
      chipHigh: 'متقدمة (>50%)',
      chipFull: '⭐ تم تحقيق 100%',
      chipTopCA: '💎 أعلى 20% مبيعات',
      showing: 'عرض',
      clientsOf: 'محل من أصل',
      btnPrint: 'طباعة / PDF',
      btnCopyWhatsApp: 'ملخص واتساب',
      thCode: 'الرمز',
      thClient: 'اسم المحل / الزبون',
      thSegment: 'الصنف',
      thCA: 'المبيعات (درهم)',
      thActivation: 'تفعيل MSL',
      thRemaining: 'المنتجات المتبقية',
      thDiagnostic: 'التشخيص',
      noResultsTitle: 'لم يتم العثور على أي محل',
      noResultsDesc: 'يرجى تغيير عبارة البحث أو إعادة ضبط التصفية.',
      btnReset: 'إعادة ضبط التصفية',
      btnPrev: 'السابق',
      btnNext: 'التالي',
      skuMatrixTitle: 'تغطية المنتجات الإلزامية (MSL)',
      skuMatrixDesc: 'تحليل نسبة انتشار ومبيعات كل منتج لدى محلات البائع المختار.',
      catHuiles: 'الزيوت',
      catEntretien: 'مواد التنظيف',
      catSoins: 'العناية والنظافة',
      catAlim: 'المواد الغذائية',
      thProduct: 'المنتج / المرجع',
      thCategory: 'الصنف',
      thBuyers: 'المحلات المشترية',
      thPenetration: 'نسبة الانتشار',
      thTotalSales: 'الحجم الإجمالي / الطناج (كغ)',
      thStatus: 'الحالة',
      leaderboardTitle: 'ترتيب بائعي الوكالة',
      leaderboardDesc: 'مقارنة الأداء الشامل في المبيعات ونسبة تفعيل المنتجات الإلزامية.',
      modalVendor: 'البائع :',
      modalCA: 'المبيعات :',
      modalActivation: 'كوتا التفعيل الإلزامي (MSL)',
      modalRemaining: 'المنتجات الناقصة لتأهيل المحل لخصم 2%',
      modalPillarsTitle: 'تأهيل المحل لخصم MSL الشهري (2% للبقال): الأركان الأساسية الـ 7 وشرط MTO',
      modalPitchTitle: 'تشخيص المحل وأهم أولويات البيع لإقناع البقال',
      modalPitchTag: 'أولويات الزيارة',
      modalOrderedSkus: 'المنتجات المطلوبة من الكتالوج (الحجم بالكيلوغرام)',
      modalMissingSkus: 'المنتجات غير المطلوبة من الكتالوج',
      modalSearchOrderedPlaceholder: 'ابحث في المواد المطلوبة...',
      modalSearchMissingPlaceholder: 'ابحث في المواد غير المطلوبة...',
      modalNoSkuFound: 'لم يتم العثور على أي مادة',
      modalBtnWhatsApp: 'مشاركة عبر واتساب',
      modalBtnClose: 'إغلاق',
      searchPlaceholder: 'ابحث باسم المحل أو رمز الزبون...',
      skuSearchPlaceholder: 'ابحث عن منتج...',
    },
  };

  const featureExplanations = {
    vendor_selector: {
      fr: {
        title: 'Filtrage Vendeur',
        desc: "Permet d'isoler uniquement votre tournée parmi tous les vendeurs de l'agence. Cliquez sur l'étoile pour mémoriser votre profil à chaque ouverture.",
      },
      ar: {
        title: 'اختيار البائع',
        desc: 'يتيح لك عزل محلاتك وزبنائك فقط من بين جميع بائعي الوكالة. اضغط على النجمة لحفظ حسابك دائماً في المرات القادمة.',
      },
    },
    ca_ttc: {
      fr: {
        title: "Chiffre d'Affaires Réalisé",
        desc: 'Montant total TTC généré sur votre tournée pour la période en cours. Il reflète votre volume de vente global.',
      },
      ar: {
        title: 'رقم المعاملات المحقق',
        desc: 'المجموع الكلي لمبيعات محلاتك بالدرهم (شامل للضريبة) خلال هذه الفترة. يمثل حجم تجارتك الكلي.',
      },
    },
    msl_rate: {
      fr: {
        title: "Taux d'Activation MSL",
        desc: 'Pourcentage des références obligatoires (Must Stock List) achetées par vos clients. Plus ce taux est élevé, plus votre prime augmente.',
      },
      ar: {
        title: 'نسبة التفعيل الإلزامي',
        desc: 'نسبة المنتجات الإلزامية التي اشترتها المحلات مقارنة بالهدف المحدد. تحدد جودة التوزيع وقيمة المكافأة.',
      },
    },
    portfolio: {
      fr: {
        title: 'Portefeuille Clients',
        desc: 'Nombre total de boutiques dans votre secteur. Les clients dormants (0 DH) nécessitent une relance prioritaire.',
      },
      ar: {
        title: 'محفظة الزبناء',
        desc: 'إجمالي المحلات في جولتك. المحلات النشطة اشترت منتجاً، والمحلات الراكدة (0 درهم) تحتاج زيارة وتنشيط.',
      },
    },
    quick_wins: {
      fr: {
        title: 'Opportunités Remise MSL (2% Magasin)',
        desc: "Magasins à qui il ne manque que 1 à 3 références pour valider 100% de leur quota MSL et débloquer leur Remise de 2% sur leur facture mensuelle. Priorité absolue pour vos visites afin de faire gagner au client sa remise !",
      },
      ar: {
        title: 'فرص تأهيل المحلات لخصم 2%',
        desc: 'محلات ينقصها من 1 إلى 3 منتجات فقط للوصول إلى هدف 100% وتمكين صاحب المحل من الاستفادة من خصم الـ MSL التجاري (2% على كامل فاتورته الشهرية). أولوية قصوى لزياراتك لإقناع البقال بطلب المواد المتبقية لصالحه.',
      },
    },
    segments: {
      fr: {
        title: 'Segments et Quotas',
        desc: 'Classification par taille : A (40 articles), B (30 articles), C (20 articles), D (10 articles). Définit le quota attendu.',
      },
      ar: {
        title: 'أصناف المحلات والأهداف',
        desc: 'تصنيف حسب حجم المحل: أ (40 منتج)، ب (30 منتج)، ج (20 منتج)، د (10 منتجات). يحدد عدد المنتجات الواجب تواجدها بالمحل.',
      },
    },
    diagnostic: {
      fr: {
        title: 'Diagnostic & Pitch',
        desc: 'Affiche la liste exacte des produits manquants les plus demandés à proposer à ce client lors de votre visite.',
      },
      ar: {
        title: 'تشخيص المحل وفرص البيع',
        desc: 'يعرض قائمة المنتجات الناقصة الأكثر طلباً في منطقته لتقترحها عليه مباشرة أثناء الزيارة.',
      },
    },
    export: {
      fr: {
        title: 'Export Excel Vendeur',
        desc: 'Télécharge un fichier Excel propre avec seulement vos clients : une feuille de synthèse et la matrice complète.',
      },
      ar: {
        title: 'تصدير ملف إكسيل للبائع',
        desc: 'تنزيل ملف إكسيل يحتوي فقط على محلاتك: ورقة ملخصة وجدول كامل لتقديمه لمسؤولك.',
      },
    },
    sku_matrix: {
      fr: {
        title: 'Matrice de Pénétration MSL',
        desc: "Affiche pour chaque article du catalogue son taux de présence chez vos clients et son volume en Kilogrammes (Tonnage), avec un filtre par famille de produits pour détecter les manques.",
      },
      ar: {
        title: 'مصفوفة انتشار المنتجات (MSL Matrix)',
        desc: 'توضح نسبة تواجد ومبيعات كل منتج في الكتالوج لدى محلاتك بالحجم والكيلوغرام (Tonnage/Kg)، مع إمكانية الفرز حسب صنف المنتجات.',
      },
    },
    leaderboard: {
      fr: {
        title: "Classement des Vendeurs de l'Agence",
        desc: "Comparatif de performance globale des vendeurs de l'agence sur le chiffre d'affaires et le taux d'activation MSL pour situer votre progression.",
      },
      ar: {
        title: 'ترتيب وتنافس بائعي الوكالة',
        desc: 'مقارنة شاملة لجميع بائعي الوكالة في رقم المعاملات ونسبة تفعيل المنتجات الإلزامية لمعرفة ترتيبك والمنافسة على المراتب الأولى.',
      },
    },
    modal_msl_score: {
      fr: {
        title: 'Quota MSL vs Total Commandé',
        desc: "Pourquoi 22 / 40 alors que 34 articles sont commandés ? Le catalogue global compte 108 articles. Pour ce magasin de segment A, l'entreprise fixe un quota obligatoire (MSL) de 40 articles stratégiques. Ce client a acheté 34 articles au total : 22 d'entre eux font partie du quota MSL obligatoire (soit 55.0%), et 12 sont des articles additionnels hors quota. Il reste donc 18 références obligatoires à lui vendre pour valider les 100%.",
      },
      ar: {
        title: 'الفرق بين كوتا MSL (22/40) وإجمالي المشتريات (34)',
        desc: "لماذا 22 من 40 بينما اشترى المحل 34 منتجاً ؟ لأن الكتالوج العام للشركة يضم 108 منتجاً. ولمحلات الصنف (أ)، حددت الشركة كوتا إلزامية (MSL) من 40 منتجاً استراتيجياً. هذا المحل اشترى 34 منتجاً في المجموع: منها 22 منتجاً محتسبة ضمن الكوتا الإلزامية (55%)، و 12 منتجاً إضافية خارج الكوتا تزيد المبيعات لكن لا تحتسب في الكوتا. وللوصول إلى 100%، يحتاج المحل لشراء الـ 18 منتجاً إلزامياً المتبقية.",
      },
    },
    pillars_eligibility: {
      fr: {
        title: 'Les 7 Piliers Incontournables & Règle MTO (Remise MSL 2%)',
        desc: "Pour que le magasin bénéficie de la Remise MSL de 2% sur sa facture du mois, il ne suffit pas d'atteindre le nombre d'articles requis : il doit obligatoirement valider chacun des 7 piliers essentiels en respectant le seuil de commande minimale officiel (MTO) du même calibre :\n" +
          "1. HT (Huiles de Table) : ≥ 6 Litres (même calibre). Si le client achète 5L, cela ne valide pas le pilier car le minimum est 6L !\n" +
          "2. HO (Huile d'Olive) : ≥ 3 Litres (même calibre)\n" +
          "3. Familia : ≥ 1 Kg (même calibre)\n" +
          "4. LEDDA : ≥ 1 Kg (même calibre)\n" +
          "5. Taous : ≥ 900 g (même calibre et même parfum)\n" +
          "6. El Ekef : ≥ 2,7 Kg (même calibre et même parfum)\n" +
          "7. El Menjel : ≥ 2,7 Kg (même calibre)\n" +
          "Ces 7 piliers validés permettent à l'épicier de toucher ses 2% de remise sur tout son chiffre d'affaires du mois !",
      },
      ar: {
        title: 'الأركان الـ 7 الأساسية وقاعدة شرط الحد الأدنى MTO (خصم 2%)',
        desc: "لكي يستفيد صاحب المحل (البقال) من خصم MSL البالغ 2% على كامل فاتورته الشهرية، تشترط الشركة إلزامياً تفعيل مادة واحدة على الأقل في كل ركن من الأركان الاستراتيجية الـ 7 وبلوغ شرط الحد الأدنى الرسمي للطلبية (MTO - Minimum To Order) من نفس الصنف :\n" +
          "1. زيوت المائدة (HT) : 6 لتر على الأقل من نفس الصنف (مثلاً: شراء 5 لتر فقط لا يحتسب ولا يؤهل الركن، فالحد الأدنى هو 6 لتر).\n" +
          "2. زيت الزيتون (HO) : 3 لتر على الأقل من نفس الصنف.\n" +
          "3. فاميليا (Familia) : 1 كغ على الأقل من نفس الصنف.\n" +
          "4. ليدا (LEDDA) : 1 كغ على الأقل من نفس الصنف.\n" +
          "5. طاووس (Taous) : 900 غ على الأقل من نفس الصنف والعبير.\n" +
          "6. الكف (El Ekef) : 2.7 كغ على الأقل من نفس الصنف والعبير.\n" +
          "7. المنجل (El Menjel) : 2.7 كغ على الأقل من نفس الصنف.\n" +
          "تحقيق هذه الأركان الـ 7 مع كوتا الصنف هو الذي يمنح البقال خصم الـ 2%، وهو أقوى حجة تفاوضية للبائع !",
      },
    },
    modal_remaining_score: {
      fr: {
        title: 'Références Restantes pour la Remise 2% du Client',
        desc: "Nombre d'articles prioritaires (MSL) restant à commander par ce magasin pour valider 100% de son quota et débloquer sa Remise MSL de 2% sur sa facture du mois.",
      },
      ar: {
        title: 'المواد المتبقية لتمكين المحل من خصم 2%',
        desc: 'عدد المواد الاستراتيجية المتبقية لهذا المحل ليحقق 100% من الكوتا المطلوبة، مما يُكسب صاحب المحل خصم MSL بنسبة 2% على مجموع مشترياته الشهرية.',
      },
    },
    modal_bought_skus: {
      fr: {
        title: 'Articles Commandés au Catalogue (Tonnage en Kg)',
        desc: 'Totalité des articles commandés par ce magasin dans le catalogue avec les volumes livrés en Kilogrammes (Tonnage). 22 articles valident directement le quota MSL et 12 articles additionnels apportent du tonnage complémentaire hors quota.',
      },
      ar: {
        title: 'المنتجات المطلوبة من الكتالوج (الحجم بالكيلوغرام / الطناج)',
        desc: 'جميع المواد التي طلبها هذا المحل من الكتالوج العام مع الكميات المباعة بالكيلوغرام / الطناج (Tonnage/Kg). 22 مادة منها محتسبة في كوتا MSL و 12 مادة إضافية خارج الكوتا تساهم في رفع المبيعات الإجمالية.',
      },
    },
    modal_missing_skus: {
      fr: {
        title: 'Articles Non Détenus au Catalogue (74)',
        desc: 'Articles du catalogue non commandés par ce client (74 sur 108). Seules 18 références spécifiques sont obligatoires pour valider les 100% MSL, le reste correspond à des ventes complémentaires optionnelles.',
      },
      ar: {
        title: 'المنتجات غير المطلوبة من الكتالوج (74)',
        desc: 'المواد الـ 74 التي لم يطلبها هذا المحل من أصل 108 بالكتالوج. 18 مادة فقط منها هي المطلوبة إلزامياً لإكمال هدف الـ 100%، والـ 56 الباقية اختيارية.',
      },
    },
    modal_pitch: {
      fr: {
        title: 'Articles Recommandés en Priorité',
        desc: 'Ces articles sont les plus vendus chez les autres boutiques du même segment. Les proposer en premier permet de combler le quota MSL restant le plus rapidement possible.',
      },
      ar: {
        title: 'أولويات البيع المقترحة',
        desc: 'هذه أكثر المنتجات طلباً لدى باقي المحلات من نفس الصنف والتي لا توجد لدى هذا المحل. اقتراحها أولاً هو أسرع وسيلة لإكمال كوتا الـ 18 منتجاً المتبقية.',
      },
    },
  };

  // =========================================================================
  // APPLICATION STATE
  // =========================================================================
  const state = {
    records: [],
    allSkus: [],
    fileName: '',
    vendors: [],
    activeVendor: 'ALL',
    pinnedVendor: localStorage.getItem('vento_pinned_vendor') || null,
    activeTab: 'clients',
    theme: localStorage.getItem('vento_theme') || 'dark',
    lang: localStorage.getItem('vento_lang') || 'fr',

    // Client tab filters
    searchQuery: '',
    segmentFilter: 'ALL',
    smartFilter: 'ALL',
    sortKey: 'ca_desc',
    currentPage: 1,
    pageSize: 35,

    // Sku tab filters
    skuSearchQuery: '',
    skuCategoryFilter: 'ALL',

    // Modal SKU search queries
    modalActiveSearchQuery: '',
    modalMissingSearchQuery: '',

    selectedClientForModal: null,
    activePopover: null,
  };

  // =========================================================================
  // DOM ELEMENT REFERENCES
  // =========================================================================
  const dom = {
    themeToggle: document.getElementById('themeToggle'),
    themeIconSun: document.getElementById('themeIconSun'),
    themeIconMoon: document.getElementById('themeIconMoon'),
    langToggle: document.getElementById('langToggle'),
    langToggleText: document.getElementById('langToggleText'),
    btnInstallPwa: document.getElementById('btnInstallPwa'),

    dropZoneSection: document.getElementById('dropZoneSection'),
    dropZoneArea: document.getElementById('dropZoneArea'),
    excelFileInput: document.getElementById('excelFileInput'),
    excelFileInputDrop: document.getElementById('excelFileInputDrop'),

    mainWorkspace: document.getElementById('mainWorkspace'),
    activeFileIndicator: document.getElementById('activeFileIndicator'),
    activeFileName: document.getElementById('activeFileName'),
    activeFileStats: document.getElementById('activeFileStats'),
    btnExportVendorExcel: document.getElementById('btnExportVendorExcel'),

    vendorSelect: document.getElementById('vendorSelect'),
    btnSetDefaultVendor: document.getElementById('btnSetDefaultVendor'),
    starBtnText: document.getElementById('starBtnText'),
    vendorPillList: document.getElementById('vendorPillList'),

    // View tabs
    tabBtns: document.querySelectorAll('.tab-btn'),
    tabPaneClients: document.getElementById('tabContentClients'),
    tabPaneProducts: document.getElementById('tabContentProducts'),
    tabPaneLeaderboard: document.getElementById('tabContentLeaderboard'),

    // Executive KPIs
    kpiTotalCA: document.getElementById('kpiTotalCA'),
    kpiAvgBasket: document.getElementById('kpiAvgBasket'),
    kpiActivationRate: document.getElementById('kpiActivationRate'),
    kpiSkuCoverage: document.getElementById('kpiSkuCoverage'),
    kpiTotalClients: document.getElementById('kpiTotalClients'),
    kpiActiveClientsBadge: document.getElementById('kpiActiveClientsBadge'),
    kpiInactiveClients: document.getElementById('kpiInactiveClients'),
    kpi100PercentClients: document.getElementById('kpi100PercentClients'),
    kpiQuickWinsCount: document.getElementById('kpiQuickWinsCount'),

    // Client tab controls
    clientSearchInput: document.getElementById('clientSearchInput'),
    btnClearSearch: document.getElementById('btnClearSearch'),
    segmentFilterPills: document.querySelectorAll('#segmentFilterPills .pill-btn'),
    sortSelect: document.getElementById('sortSelect'),
    smartChips: document.querySelectorAll('#smartChipsGroup .chip-btn'),
    visibleClientsCount: document.getElementById('visibleClientsCount'),
    totalFilterClientsCount: document.getElementById('totalFilterClientsCount'),
    clientsTableBody: document.getElementById('clientsTableBody'),
    noResultsState: document.getElementById('noResultsState'),
    btnResetFilters: document.getElementById('btnResetFilters'),
    paginationInfo: document.getElementById('paginationInfo'),
    btnPrevPage: document.getElementById('btnPrevPage'),
    btnNextPage: document.getElementById('btnNextPage'),

    btnPrintRouteSheet: document.getElementById('btnPrintRouteSheet'),
    btnCopyDailySummary: document.getElementById('btnCopyDailySummary'),

    // SKU tab controls
    skuSearchInput: document.getElementById('skuSearchInput'),
    skuCategoryPills: document.querySelectorAll('#skuCategoryPills .pill-btn'),
    skusTableBody: document.getElementById('skusTableBody'),

    // Leaderboard
    leaderboardGrid: document.getElementById('leaderboardGrid'),
    agencyBadgeName: document.getElementById('agencyBadgeName'),

    // Client Modal
    clientModal: document.getElementById('clientModal'),
    btnCloseModal: document.getElementById('btnCloseModal'),
    btnCloseModalBottom: document.getElementById('btnCloseModalBottom'),
    modalClientSegBadge: document.getElementById('modalClientSegBadge'),
    modalClientCode: document.getElementById('modalClientCode'),
    modalClientName: document.getElementById('modalClientName'),
    modalClientVendor: document.getElementById('modalClientVendor'),
    modalClientCA: document.getElementById('modalClientCA'),
    modalScoreText: document.getElementById('modalScoreText'),
    modalCatalogBreakdownText: document.getElementById('modalCatalogBreakdownText'),
    modalRemainingText: document.getElementById('modalRemainingText'),
    modalRemainingSubText: document.getElementById('modalRemainingSubText'),
    modalProgressBar: document.getElementById('modalProgressBar'),
    modalPillarsStatusBadge: document.getElementById('modalPillarsStatusBadge'),
    modalPillarsGrid: document.getElementById('modalPillarsGrid'),
    modalPitchRecommendations: document.getElementById('modalPitchRecommendations'),
    modalActiveSkusCount: document.getElementById('modalActiveSkusCount'),
    modalActiveSkusSub: document.getElementById('modalActiveSkusSub'),
    modalActiveSkusSearch: document.getElementById('modalActiveSkusSearch'),
    modalActiveSkusList: document.getElementById('modalActiveSkusList'),
    modalMissingSkusCount: document.getElementById('modalMissingSkusCount'),
    modalMissingSkusSub: document.getElementById('modalMissingSkusSub'),
    modalMissingSkusSearch: document.getElementById('modalMissingSkusSearch'),
    modalMissingSkusList: document.getElementById('modalMissingSkusList'),
    btnShareWhatsAppClient: document.getElementById('btnShareWhatsAppClient'),

    // Feature Explanations Popover
    infoPopover: document.getElementById('infoPopover'),
    popoverTitle: document.getElementById('popoverTitle'),
    popoverContent: document.getElementById('popoverContent'),
    btnClosePopover: document.getElementById('btnClosePopover'),

    toastContainer: document.getElementById('toastContainer'),
  };

  // =========================================================================
  // UTILITIES & FORMATTERS
  // =========================================================================
  const fmt = {
    currency: (n) => {
      const num = Number(n) || 0;
      const unit = state.lang === 'ar' ? ' درهم' : ' DH';
      return num.toLocaleString('fr-FR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + unit;
    },
    tonnage: (n) => {
      const num = Number(n) || 0;
      if (num >= 1000) {
        const tonnes = num / 1000;
        const unit = state.lang === 'ar' ? ' طن' : ' T';
        return tonnes.toLocaleString('fr-FR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + unit;
      }
      const unit = state.lang === 'ar' ? ' كغ' : ' Kg';
      return num.toLocaleString('fr-FR', { minimumFractionDigits: num % 1 === 0 ? 0 : 2, maximumFractionDigits: 2 }) + unit;
    },
    int: (n) => {
      const num = Number(n) || 0;
      return num.toLocaleString('fr-FR');
    },
    pct: (n) => {
      const num = Number(n) || 0;
      return num.toFixed(1) + '%';
    },
    stripAccents: (str) => {
      if (!str) return '';
      return str
        .toString()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .trim();
    },
  };

  // =========================================================================
  // 7 ESSENTIAL ELIGIBILITY PILLARS & MTO MECHANISM (October 2026 Direct Deployment Rules)
  // =========================================================================
  const SEVEN_PILLARS = [
    {
      id: 'HT',
      brand: 'HT',
      nameFr: 'Huiles de Table',
      nameAr: 'زيوت المائدة',
      icon: '🛢️',
      mto: 6.0,
      mtoDisplayFr: '6 L',
      mtoDisplayAr: '6 لتر',
      ruleFr: 'Même calibre (≥ 6 L requis)',
      ruleAr: 'نفس الصنف (6 لتر على الأقل)',
      matches: (s) => /CRISTAL|HUILOR|MABROUKA|\b3G\b|FRITURE|H\.T/i.test(s) && !/MAY|MOUTARD/i.test(s),
    },
    {
      id: 'HO',
      brand: 'HO',
      nameFr: "Huile d'Olive",
      nameAr: 'زيت الزيتون',
      icon: '🫒',
      mto: 3.0,
      mtoDisplayFr: '3 L',
      mtoDisplayAr: '3 لتر',
      ruleFr: 'Même calibre (≥ 3 L requis)',
      ruleAr: 'نفس الصنف (3 لتر على الأقل)',
      matches: (s) => /ALHORRA/i.test(s),
    },
    {
      id: 'Familia',
      brand: 'Familia',
      nameFr: 'Familia',
      nameAr: 'فاميليا',
      icon: '🧈',
      mto: 1.0,
      mtoDisplayFr: '1 Kg',
      mtoDisplayAr: '1 كغ',
      ruleFr: 'Même calibre (≥ 1 Kg requis)',
      ruleAr: 'نفس الصنف (1 كغ على الأقل)',
      matches: (s) => /FAMILIA/i.test(s),
    },
    {
      id: 'LEDDA',
      brand: 'LEDDA',
      nameFr: 'LEDDA',
      nameAr: 'ليدا',
      icon: '🧈',
      mto: 1.0,
      mtoDisplayFr: '1 Kg',
      mtoDisplayAr: '1 كغ',
      ruleFr: 'Même calibre (≥ 1 Kg requis)',
      ruleAr: 'نفس الصنف (1 كغ على الأقل)',
      matches: (s) => /LEDDA/i.test(s),
    },
    {
      id: 'Taous',
      brand: 'Taous',
      nameFr: 'Taous',
      nameAr: 'طاووس',
      icon: '🧼',
      mto: 0.900,
      mtoDisplayFr: '900 g',
      mtoDisplayAr: '900 غ',
      ruleFr: 'Même calibre et parfum (≥ 900 g requis)',
      ruleAr: 'نفس الصنف والعبير (900 غ على الأقل)',
      matches: (s) => /TAOUS/i.test(s),
    },
    {
      id: 'El Ekef',
      brand: 'El Ekef',
      nameFr: 'El Ekef',
      nameAr: 'الكف',
      icon: '🧽',
      mto: 2.7,
      mtoDisplayFr: '2,7 Kg',
      mtoDisplayAr: '2.7 كغ',
      ruleFr: 'Même calibre et parfum (≥ 2,7 Kg requis)',
      ruleAr: 'نفس الصنف والعبير (2.7 كغ على الأقل)',
      matches: (s) => (/KEF|LIQUIDE VAISSELLE/i.test(s)) && !/MENJEL/i.test(s),
    },
    {
      id: 'El Menjel',
      brand: 'El Menjel',
      nameFr: 'El Menjel',
      nameAr: 'المنجل',
      icon: '🧼',
      mto: 2.7,
      mtoDisplayFr: '2,7 Kg',
      mtoDisplayAr: '2.7 كغ',
      ruleFr: 'Même calibre (≥ 2,7 Kg requis)',
      ruleAr: 'نفس الصنف (2.7 كغ على الأقل)',
      matches: (s) => /MENJEL/i.test(s),
    },
  ];

  function showToast(message, type = 'info', duration = 3000) {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.textContent = message;
    dom.toastContainer.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      setTimeout(() => toast.remove(), 250);
    }, duration);
  }

  // =========================================================================
  // THEME MANAGEMENT
  // =========================================================================
  function applyTheme(theme) {
    state.theme = theme;
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('vento_theme', theme);
    if (theme === 'light') {
      dom.themeIconSun.classList.add('hidden');
      dom.themeIconMoon.classList.remove('hidden');
    } else {
      dom.themeIconSun.classList.remove('hidden');
      dom.themeIconMoon.classList.add('hidden');
    }
  }

  dom.themeToggle.addEventListener('click', () => {
    applyTheme(state.theme === 'dark' ? 'light' : 'dark');
  });

  // =========================================================================
  // LOCALIZATION SYSTEM (FR / AR)
  // =========================================================================
  function applyLanguage(lang) {
    state.lang = lang;
    localStorage.setItem('vento_lang', lang);
    document.documentElement.setAttribute('lang', lang);
    document.documentElement.setAttribute('dir', lang === 'ar' ? 'rtl' : 'ltr');

    // Update button text to offer the alternative language
    dom.langToggleText.textContent = lang === 'fr' ? 'العربية' : 'Français';

    const t = i18n[lang] || i18n.fr;

    // Update text content of elements with data-i18n
    document.querySelectorAll('[data-i18n]').forEach((el) => {
      const key = el.getAttribute('data-i18n');
      if (t[key] !== undefined) {
        el.textContent = t[key];
      }
    });

    // Update placeholders
    dom.clientSearchInput.placeholder = t.searchPlaceholder;
    if (dom.skuSearchInput) dom.skuSearchInput.placeholder = t.skuSearchPlaceholder;
    if (dom.modalActiveSkusSearch) dom.modalActiveSkusSearch.placeholder = t.modalSearchOrderedPlaceholder;
    if (dom.modalMissingSkusSearch) dom.modalMissingSkusSearch.placeholder = t.modalSearchMissingPlaceholder;

    // Re-render views if dataset is loaded
    if (state.records && state.records.length > 0) {
      const clientWord = lang === 'ar' ? 'محل' : 'clients';
      const vendorWord = lang === 'ar' ? 'بائع' : 'vendeur(s)';
      dom.activeFileStats.textContent = `${state.records.length.toLocaleString('fr-FR')} ${clientWord} • ${state.vendors.length} ${vendorWord}`;
      populateVendorDropdown();
      renderVendorPills();
      renderCurrentVendorView();
      updatePinnedVendorButton();
      if (state.selectedClientForModal) {
        openClientModal(state.selectedClientForModal.code);
      }
    }
  }

  dom.langToggle.addEventListener('click', () => {
    applyLanguage(state.lang === 'fr' ? 'ar' : 'fr');
  });

  // =========================================================================
  // FEATURE EXPLANATION POPOVER (CIRCULAR ! BUTTON)
  // =========================================================================
  function setupInfoPopovers() {
    document.addEventListener('click', (e) => {
      const btn = e.target.closest('.info-pop-btn');
      if (btn) {
        e.stopPropagation();
        const key = btn.getAttribute('data-info');
        togglePopover(btn, key);
      } else if (dom.infoPopover && !dom.infoPopover.contains(e.target)) {
        closePopover();
      }
    });

    window.addEventListener('scroll', closePopover, { passive: true });
    window.addEventListener('resize', closePopover, { passive: true });

    if (dom.btnClosePopover) {
      dom.btnClosePopover.addEventListener('click', (e) => {
        e.stopPropagation();
        closePopover();
      });
    }
  }

  function togglePopover(anchorBtn, key) {
    if (state.activePopover === anchorBtn) {
      closePopover();
      return;
    }

    const exp = featureExplanations[key];
    if (!exp) return;

    const langData = exp[state.lang] || exp.fr;
    dom.popoverTitle.textContent = langData.title;
    dom.popoverContent.textContent = langData.desc;

    // Position popover relative to viewport (since position: fixed)
    const rect = anchorBtn.getBoundingClientRect();
    const isRtl = state.lang === 'ar';

    dom.infoPopover.classList.remove('hidden');
    state.activePopover = anchorBtn;

    // Calculate position
    const popRect = dom.infoPopover.getBoundingClientRect();
    let top = rect.bottom + 8;
    // If popover overflows bottom of viewport, position it above the anchor button
    if (top + popRect.height > window.innerHeight - 12) {
      top = Math.max(12, rect.top - popRect.height - 8);
    }

    let left = isRtl ? (rect.right - popRect.width) : rect.left;

    // Boundary check
    if (left < 12) left = 12;
    if (left + popRect.width > window.innerWidth - 12) {
      left = window.innerWidth - popRect.width - 12;
    }

    dom.infoPopover.style.top = `${top}px`;
    dom.infoPopover.style.left = `${left}px`;
  }

  function closePopover() {
    if (dom.infoPopover) dom.infoPopover.classList.add('hidden');
    state.activePopover = null;
  }

  // =========================================================================
  // FILE PARSING & DATA NORMALIZATION
  // =========================================================================
  function normalizeHeader(str) {
    if (!str) return '';
    return str
      .toString()
      .trim()
      .toUpperCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
  }

  function parseExcelData(binaryData, fileName = 'Fichier Suivi MSL.xlsx') {
    try {
      if (typeof XLSX === 'undefined') {
        showToast('Erreur: Bibliothèque Excel introuvable.', 'danger');
        return;
      }

      const workbook = XLSX.read(binaryData, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];
      const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

      if (!rawRows || rawRows.length < 2) {
        showToast('Fichier Excel vide ou non reconnu.', 'warning');
        return;
      }

      const headerRow = rawRows[0];
      const colMap = {};
      const skuIndices = [];

      headerRow.forEach((colName, idx) => {
        const norm = normalizeHeader(colName);
        if (norm === 'AGENCE') colMap.agence = idx;
        else if (norm === 'VENDEUR') colMap.vendeur = idx;
        else if (norm === 'CODE CLIENT' && colMap.codeClient === undefined) colMap.codeClient = idx;
        else if (norm === 'SEGMENT') colMap.segment = idx;
        else if (norm.includes('MSL') && (norm.includes('ACTIVE') || norm.includes('ACT'))) colMap.mslAct = idx;
        else if (norm.includes('TAUX') && norm.includes('ACT')) colMap.tauxAct = idx;
        else if (norm.includes('RESTANT') || norm.includes('SKU RESTANT')) colMap.skuRem = idx;
        else if (norm.includes('CA') && norm.includes('TTC')) colMap.caTTC = idx;
        else if (norm.includes('NOM CLIENT') || norm === 'NOM') colMap.nomClient = idx;
        else {
          if (idx >= 10 && colName && colName.toString().trim()) {
            skuIndices.push({ idx, name: colName.toString().trim() });
          }
        }
      });

      const records = [];
      const allSkusSet = new Set();
      skuIndices.forEach((s) => allSkusSet.add(s.name));

      for (let r = 1; r < rawRows.length; r++) {
        const row = rawRows[r];
        if (!row || row.length === 0) continue;

        const agence = colMap.agence !== undefined ? String(row[colMap.agence] || '').trim() : '';
        const vendeur = colMap.vendeur !== undefined ? String(row[colMap.vendeur] || '').trim() : '';
        const code = colMap.codeClient !== undefined ? String(row[colMap.codeClient] || '').trim() : '';
        const nom = colMap.nomClient !== undefined ? String(row[colMap.nomClient] || '').trim() : '';
        const segment = colMap.segment !== undefined ? String(row[colMap.segment] || '').trim().toUpperCase() : 'C';

        if (!code && !nom && !vendeur) continue;

        const mslAct = colMap.mslAct !== undefined ? Math.round(Number(row[colMap.mslAct]) || 0) : 0;
        let tauxAct = colMap.tauxAct !== undefined ? Number(row[colMap.tauxAct]) || 0 : 0;
        if (tauxAct <= 1.0 && tauxAct > 0) tauxAct = Math.round(tauxAct * 1000) / 10;
        const skuRem = colMap.skuRem !== undefined ? Math.round(Number(row[colMap.skuRem]) || 0) : 0;
        const caTTC = colMap.caTTC !== undefined ? Math.round((Number(row[colMap.caTTC]) || 0) * 100) / 100 : 0;

        const skus = {};
        skuIndices.forEach(({ idx, name }) => {
          const val = row[idx];
          if (val !== undefined && val !== null && val !== '') {
            const num = Number(val);
            if (!isNaN(num) && num > 0) {
              skus[name] = Math.round(num * 100) / 100;
            }
          }
        });

        records.push({
          agence: agence || 'AGENCE',
          vendeur: vendeur || 'Vendeur',
          code: code || `C-${r}`,
          segment: ['A', 'B', 'C', 'D'].includes(segment) ? segment : 'C',
          msl_act: mslAct,
          taux_act: tauxAct,
          sku_rem: skuRem,
          ca_ttc: caTTC,
          nom: nom || `Client ${code}`,
          skus: skus,
        });
      }

      loadDatasetIntoApp(records, Array.from(allSkusSet), fileName);
      showToast(state.lang === 'ar' ? `تم تحميل الملف بنجاح (${records.length} محل).` : `Fichier chargé avec succès (${records.length} clients).`, 'success');
    } catch (err) {
      console.error('Error parsing excel:', err);
      showToast('Erreur de lecture du fichier.', 'danger');
    }
  }

  function loadDatasetIntoApp(records, allSkus, fileName) {
    state.records = records;
    state.allSkus = allSkus;
    state.fileName = fileName;

    // Detect unique vendors
    const vendorMap = new Map();
    records.forEach((r) => {
      const v = r.vendeur;
      if (!vendorMap.has(v)) {
        vendorMap.set(v, { count: 0, ca: 0 });
      }
      const st = vendorMap.get(v);
      st.count++;
      st.ca += r.ca_ttc;
    });

    state.vendors = Array.from(vendorMap.keys()).sort((a, b) => vendorMap.get(b).ca - vendorMap.get(a).ca);

    if (state.pinnedVendor && state.vendors.includes(state.pinnedVendor)) {
      state.activeVendor = state.pinnedVendor;
    } else if (state.vendors.length === 1) {
      state.activeVendor = state.vendors[0];
    } else {
      state.activeVendor = state.vendors[0] || 'ALL';
    }

    // UI Updates
    dom.activeFileName.textContent = fileName;
    dom.activeFileStats.textContent = `${records.length.toLocaleString('fr-FR')} clients • ${state.vendors.length} vendeur(s)`;
    dom.activeFileIndicator.classList.remove('hidden');
    dom.btnExportVendorExcel.disabled = false;

    if (records[0] && records[0].agence) {
      dom.agencyBadgeName.textContent = `AGENCE ${records[0].agence.toUpperCase()}`;
    }

    dom.dropZoneSection.classList.add('hidden');
    dom.mainWorkspace.classList.remove('hidden');

    populateVendorDropdown();
    renderVendorPills();
    updatePinnedVendorButton();
    renderCurrentVendorView();

    savePersistentDataset({ records, allSkus, fileName });
  }

  // =========================================================================
  // PERSISTENT OFFLINE STORAGE (IndexedDB + localStorage Fallback)
  // =========================================================================
  const DB_NAME = 'VentoOfflineDB';
  const DB_STORE = 'datasets';
  const DB_KEY = 'active_dataset';

  function openOfflineDB() {
    return new Promise((resolve, reject) => {
      if (!window.indexedDB) {
        reject(new Error('IndexedDB not supported'));
        return;
      }
      const req = window.indexedDB.open(DB_NAME, 1);
      req.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains(DB_STORE)) {
          db.createObjectStore(DB_STORE);
        }
      };
      req.onsuccess = (e) => resolve(e.target.result);
      req.onerror = (e) => reject(e.target.error);
    });
  }

  function savePersistentDataset(payload) {
    // 1. Save in IndexedDB for permanent, large-capacity phone storage
    openOfflineDB()
      .then((db) => {
        const tx = db.transaction(DB_STORE, 'readwrite');
        tx.objectStore(DB_STORE).put(payload, DB_KEY);
      })
      .catch(() => {});

    // 2. Fallback to localStorage / sessionStorage
    try {
      const serialized = JSON.stringify(payload);
      localStorage.setItem('vento_persistent_dataset', serialized);
      sessionStorage.setItem('vento_session_dataset', serialized);
    } catch (e) {
      // Ignore quota errors if dataset exceeds 5MB localStorage limit (IndexedDB handles it)
    }
  }

  function loadPersistentDataset() {
    return openOfflineDB()
      .then((db) => {
        return new Promise((resolve, reject) => {
          const tx = db.transaction(DB_STORE, 'readonly');
          const req = tx.objectStore(DB_STORE).get(DB_KEY);
          req.onsuccess = () => resolve(req.result || null);
          req.onerror = () => reject(req.error);
        });
      })
      .catch(() => null)
      .then((idbData) => {
        if (idbData && idbData.records && idbData.records.length > 0) {
          return idbData;
        }
        try {
          const raw = localStorage.getItem('vento_persistent_dataset') || sessionStorage.getItem('vento_session_dataset');
          if (raw) {
            return JSON.parse(raw);
          }
        } catch (e) {
          // Ignore parse errors
        }
        return null;
      });
  }

  // =========================================================================
  // VENDOR SELECTION & PINNING LOGIC
  // =========================================================================
  function populateVendorDropdown() {
    dom.vendorSelect.innerHTML = '';
    const clientWord = state.lang === 'ar' ? 'محل' : 'clients';

    if (state.vendors.length > 1) {
      const optAll = document.createElement('option');
      optAll.value = 'ALL';
      optAll.textContent = state.lang === 'ar' 
        ? `جميع البائعين (${state.records.length} ${clientWord})` 
        : `Tous les vendeurs (${state.records.length} ${clientWord})`;
      dom.vendorSelect.appendChild(optAll);
    }

    state.vendors.forEach((v) => {
      const vRecords = state.records.filter((r) => r.vendeur === v);
      const vCA = vRecords.reduce((acc, r) => acc + r.ca_ttc, 0);
      const opt = document.createElement('option');
      opt.value = v;
      opt.textContent = state.lang === 'ar'
        ? `${v} (${vRecords.length} محل — ${fmt.currency(vCA)})`
        : `${v} (${vRecords.length} clients — ${fmt.currency(vCA)})`;
      dom.vendorSelect.appendChild(opt);
    });

    dom.vendorSelect.value = state.activeVendor;
  }

  function renderVendorPills() {
    dom.vendorPillList.innerHTML = '';
    const clientWord = state.lang === 'ar' ? 'محل' : 'clients';

    if (state.vendors.length > 1) {
      const pillAll = document.createElement('div');
      pillAll.className = `vendor-pill ${state.activeVendor === 'ALL' ? 'active' : ''}`;
      pillAll.innerHTML = `<span>${state.lang === 'ar' ? 'جميع البائعين' : 'Tous'}</span> <span class="vendor-pill-badge">${state.records.length} ${clientWord}</span>`;
      pillAll.addEventListener('click', () => switchVendor('ALL'));
      dom.vendorPillList.appendChild(pillAll);
    }

    state.vendors.forEach((v) => {
      const vRecords = state.records.filter((r) => r.vendeur === v);
      const isPinned = state.pinnedVendor === v;
      const pill = document.createElement('div');
      pill.className = `vendor-pill ${state.activeVendor === v ? 'active' : ''}`;
      pill.innerHTML = `
        <span>${isPinned ? '⭐ ' : ''}${v}</span>
        <span class="vendor-pill-badge">${vRecords.length} ${clientWord}</span>
      `;
      pill.addEventListener('click', () => switchVendor(v));
      dom.vendorPillList.appendChild(pill);
    });
  }

  function switchVendor(vendorName) {
    state.activeVendor = vendorName;
    dom.vendorSelect.value = vendorName;
    state.currentPage = 1;
    renderVendorPills();
    updatePinnedVendorButton();
    renderCurrentVendorView();
  }

  dom.vendorSelect.addEventListener('change', (e) => {
    switchVendor(e.target.value);
  });

  function updatePinnedVendorButton() {
    if (state.activeVendor === 'ALL') {
      dom.btnSetDefaultVendor.classList.add('hidden');
      return;
    }
    dom.btnSetDefaultVendor.classList.remove('hidden');

    const isPinned = state.pinnedVendor === state.activeVendor;
    const t = i18n[state.lang] || i18n.fr;

    if (isPinned) {
      dom.btnSetDefaultVendor.classList.add('pinned');
      dom.starBtnText.textContent = t.pinnedProfile;
    } else {
      dom.btnSetDefaultVendor.classList.remove('pinned');
      dom.starBtnText.textContent = t.pinProfile;
    }
  }

  dom.btnSetDefaultVendor.addEventListener('click', () => {
    if (state.activeVendor === 'ALL') return;

    if (state.pinnedVendor === state.activeVendor) {
      state.pinnedVendor = null;
      localStorage.removeItem('vento_pinned_vendor');
      showToast(state.lang === 'ar' ? 'تم إلغاء الحساب الافتراضي.' : 'Profil par défaut effacé.', 'info');
    } else {
      state.pinnedVendor = state.activeVendor;
      localStorage.setItem('vento_pinned_vendor', state.activeVendor);
      showToast(state.lang === 'ar' ? `تم حفظ "${state.activeVendor}" كحسابك الافتراضي!` : `"${state.activeVendor}" est mémorisé comme votre profil!`, 'success');
    }
    renderVendorPills();
    updatePinnedVendorButton();
  });

  // =========================================================================
  // FILTERING & AGGREGATION LOGIC
  // =========================================================================
  function getVendorRecords() {
    if (state.activeVendor === 'ALL') {
      return state.records;
    }
    return state.records.filter((r) => r.vendeur === state.activeVendor);
  }

  function computeKPIs(vendorRecords) {
    let totalCA = 0;
    let activeClients = 0;
    let mslActSum = 0;
    let mslTargetSum = 0;
    let clients100 = 0;
    let quickWins = 0;

    vendorRecords.forEach((r) => {
      totalCA += r.ca_ttc;
      const target = r.msl_act + r.sku_rem;
      mslActSum += r.msl_act;
      mslTargetSum += target;

      if (r.ca_ttc > 0 || r.msl_act > 0) activeClients++;
      if (r.sku_rem === 0 && r.msl_act > 0) clients100++;
      if (r.sku_rem > 0 && r.sku_rem <= 3) quickWins++;
    });

    const totalClients = vendorRecords.length;
    const inactiveClients = totalClients - activeClients;
    const avgBasket = activeClients > 0 ? totalCA / activeClients : 0;
    const activationRate = mslTargetSum > 0 ? (mslActSum / mslTargetSum) * 100 : 0;

    return {
      totalCA,
      totalClients,
      activeClients,
      inactiveClients,
      avgBasket,
      activationRate,
      mslActSum,
      mslTargetSum,
      clients100,
      quickWins,
    };
  }

  function updateKPIDisplay(kpis) {
    dom.kpiTotalCA.textContent = fmt.currency(kpis.totalCA);
    dom.kpiAvgBasket.textContent = fmt.currency(kpis.avgBasket);

    dom.kpiActivationRate.textContent = fmt.pct(kpis.activationRate);
    dom.kpiSkuCoverage.textContent = `${fmt.int(kpis.mslActSum)} / ${fmt.int(kpis.mslTargetSum)} SKUs`;

    dom.kpiTotalClients.textContent = fmt.int(kpis.totalClients);
    dom.kpiActiveClientsBadge.textContent = `${fmt.int(kpis.activeClients)} ${state.lang === 'ar' ? 'نشط' : 'actifs'}`;
    dom.kpiInactiveClients.textContent = fmt.int(kpis.inactiveClients);

    dom.kpi100PercentClients.textContent = fmt.int(kpis.clients100);
    dom.kpiQuickWinsCount.textContent = fmt.int(kpis.quickWins);
  }

  function updateSmartFilterCounts(vendorRecords) {
    const counts = {
      ALL: vendorRecords.length,
      QUICK_WINS: 0,
      ZERO_ACT: 0,
      MID_ACT: 0,
      HIGH_ACT: 0,
      FULL_100: 0,
      TOP_CA: 0,
    };

    const sortedCA = vendorRecords.map((r) => r.ca_ttc).sort((a, b) => b - a);
    const topThresholdIndex = Math.floor(sortedCA.length * 0.2);
    const caThreshold = sortedCA[topThresholdIndex] || 0;

    vendorRecords.forEach((r) => {
      if (r.sku_rem > 0 && r.sku_rem <= 3) counts.QUICK_WINS++;
      if (r.msl_act === 0) counts.ZERO_ACT++;
      if (r.taux_act > 0 && r.taux_act <= 50) counts.MID_ACT++;
      if (r.taux_act > 50 && r.sku_rem > 0) counts.HIGH_ACT++;
      if (r.sku_rem === 0 && r.msl_act > 0) counts.FULL_100++;
      if (r.ca_ttc >= caThreshold && r.ca_ttc > 0) counts.TOP_CA++;
    });

    document.getElementById('chipCountAll').textContent = counts.ALL;
    document.getElementById('chipCountQuickWins').textContent = counts.QUICK_WINS;
    document.getElementById('chipCountZero').textContent = counts.ZERO_ACT;
    document.getElementById('chipCountMid').textContent = counts.MID_ACT;
    document.getElementById('chipCountHigh').textContent = counts.HIGH_ACT;
    document.getElementById('chipCountFull').textContent = counts.FULL_100;
    document.getElementById('chipCountTopCA').textContent = counts.TOP_CA;
  }

  function getFilteredClients(vendorRecords) {
    let list = [...vendorRecords];

    if (state.segmentFilter !== 'ALL') {
      list = list.filter((r) => r.segment === state.segmentFilter);
    }

    if (state.smartFilter !== 'ALL') {
      if (state.smartFilter === 'QUICK_WINS') {
        list = list.filter((r) => r.sku_rem > 0 && r.sku_rem <= 3);
      } else if (state.smartFilter === 'ZERO_ACT') {
        list = list.filter((r) => r.msl_act === 0);
      } else if (state.smartFilter === 'MID_ACT') {
        list = list.filter((r) => r.taux_act > 0 && r.taux_act <= 50);
      } else if (state.smartFilter === 'HIGH_ACT') {
        list = list.filter((r) => r.taux_act > 50 && r.sku_rem > 0);
      } else if (state.smartFilter === 'FULL_100') {
        list = list.filter((r) => r.sku_rem === 0 && r.msl_act > 0);
      } else if (state.smartFilter === 'TOP_CA') {
        const sorted = vendorRecords.map((r) => r.ca_ttc).sort((a, b) => b - a);
        const threshold = sorted[Math.floor(sorted.length * 0.2)] || 0;
        list = list.filter((r) => r.ca_ttc >= threshold && r.ca_ttc > 0);
      }
    }

    if (state.searchQuery) {
      const q = fmt.stripAccents(state.searchQuery);
      list = list.filter((r) => fmt.stripAccents(r.nom).includes(q) || fmt.stripAccents(r.code).includes(q));
    }

    list.sort((a, b) => {
      switch (state.sortKey) {
        case 'ca_desc': return b.ca_ttc - a.ca_ttc;
        case 'ca_asc': return a.ca_ttc - b.ca_ttc;
        case 'act_desc': return b.taux_act - a.taux_act;
        case 'act_asc': return a.taux_act - b.taux_act;
        case 'rem_asc': return a.sku_rem - b.sku_rem;
        case 'nom_asc': return a.nom.localeCompare(b.nom);
        default: return b.ca_ttc - a.ca_ttc;
      }
    });

    return list;
  }

  // =========================================================================
  // RENDER: CLIENTS TABLE & PAGINATION
  // =========================================================================
  function renderClientsTable() {
    const vendorRecords = getVendorRecords();
    const filtered = getFilteredClients(vendorRecords);

    dom.totalFilterClientsCount.textContent = vendorRecords.length;
    dom.visibleClientsCount.textContent = filtered.length;

    if (filtered.length === 0) {
      dom.clientsTableBody.innerHTML = '';
      dom.noResultsState.classList.remove('hidden');
      dom.paginationInfo.textContent = state.lang === 'ar' ? 'صفحة 0 من 0' : 'Page 0 sur 0';
      dom.btnPrevPage.disabled = true;
      dom.btnNextPage.disabled = true;
      return;
    }

    dom.noResultsState.classList.add('hidden');

    const totalPages = Math.ceil(filtered.length / state.pageSize);
    if (state.currentPage > totalPages) state.currentPage = totalPages;
    if (state.currentPage < 1) state.currentPage = 1;

    const startIdx = (state.currentPage - 1) * state.pageSize;
    const pageRecords = filtered.slice(startIdx, startIdx + state.pageSize);

    dom.paginationInfo.textContent = state.lang === 'ar' 
      ? `صفحة ${state.currentPage} من ${totalPages} (${filtered.length} محل)`
      : `Page ${state.currentPage} sur ${totalPages} (${filtered.length} clients)`;
    dom.btnPrevPage.disabled = state.currentPage <= 1;
    dom.btnNextPage.disabled = state.currentPage >= totalPages;

    dom.clientsTableBody.innerHTML = pageRecords
      .map((client) => {
        const isQuickWin = client.sku_rem > 0 && client.sku_rem <= 3;
        const isFull100 = client.sku_rem === 0 && client.msl_act > 0;
        const target = client.msl_act + client.sku_rem;

        let barClass = 'bar-gray';
        if (client.taux_act >= 100) barClass = 'bar-green';
        else if (client.taux_act > 50) barClass = 'bar-blue';
        else if (client.taux_act > 0) barClass = 'bar-amber';

        let remBadgeClass = 'rem-badge-normal';
        let remText = state.lang === 'ar' ? `${client.sku_rem} متبقي` : `${client.sku_rem} restants`;
        if (isFull100) {
          remBadgeClass = 'rem-badge-zero';
          remText = state.lang === 'ar' ? '✓ تم الهدف' : '✓ Validé';
        } else if (isQuickWin) {
          remBadgeClass = 'rem-badge-few';
        }

        const quickWinBadgeText = state.lang === 'ar' ? '🚀 قريب من 100%' : '🚀 Proche 100%';
        const diagnosticBtnText = state.lang === 'ar' ? 'تشخيص' : 'Diagnostic';

        return `
        <tr data-client-code="${client.code}">
          <td class="client-code-cell">${client.code}</td>
          <td>
            <div class="client-name-cell">
              <span>${client.nom}</span>
              ${isQuickWin ? `<span class="client-quick-badge">${quickWinBadgeText}</span>` : ''}
              ${isFull100 ? `<span class="text-success font-mono" style="font-size:0.7rem">⭐ 100%</span>` : ''}
            </div>
          </td>
          <td>
            <span class="badge-seg badge-seg-${client.segment.toLowerCase()}">SEG ${client.segment}</span>
          </td>
          <td class="text-right font-mono" style="font-weight:700">
            ${fmt.currency(client.ca_ttc)}
          </td>
          <td>
            <div class="table-act-wrap">
              <div class="table-act-header">
                <span class="table-act-rate ${client.taux_act >= 100 ? 'text-success' : ''}">${client.taux_act.toFixed(1)}%</span>
                <span class="table-act-skus">${client.msl_act} / ${target}</span>
              </div>
              <div class="table-progress-track">
                <div class="table-progress-bar ${barClass}" style="width: ${Math.min(100, client.taux_act)}%"></div>
              </div>
            </div>
          </td>
          <td class="text-center">
            <span class="rem-badge ${remBadgeClass}">
              ${remText}
            </span>
          </td>
          <td class="text-center">
            <button class="btn btn-secondary btn-sm btn-open-modal" data-code="${client.code}" title="Voir diagnostic">
              <span>${diagnosticBtnText}</span>
            </button>
          </td>
        </tr>
      `;
      })
      .join('');

    // Attach click handlers
    dom.clientsTableBody.querySelectorAll('.btn-open-modal').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        openClientModal(btn.getAttribute('data-code'));
      });
    });

    // Row click
    dom.clientsTableBody.querySelectorAll('tr').forEach((tr) => {
      tr.style.cursor = 'pointer';
      tr.addEventListener('click', (e) => {
        if (e.target.closest('button') || e.target.closest('input')) return;
        const code = tr.getAttribute('data-client-code');
        if (code) openClientModal(code);
      });
    });
  }

  // Client Search & Clear
  dom.clientSearchInput.addEventListener('input', (e) => {
    state.searchQuery = e.target.value;
    if (state.searchQuery) dom.btnClearSearch.classList.remove('hidden');
    else dom.btnClearSearch.classList.add('hidden');
    state.currentPage = 1;
    renderClientsTable();
  });

  dom.btnClearSearch.addEventListener('click', () => {
    dom.clientSearchInput.value = '';
    state.searchQuery = '';
    dom.btnClearSearch.classList.add('hidden');
    state.currentPage = 1;
    renderClientsTable();
  });

  // Segment Filter Pills
  dom.segmentFilterPills.forEach((pill) => {
    pill.addEventListener('click', () => {
      dom.segmentFilterPills.forEach((p) => p.classList.remove('active'));
      pill.classList.add('active');
      state.segmentFilter = pill.getAttribute('data-seg');
      state.currentPage = 1;
      renderClientsTable();
    });
  });

  // Sort dropdown
  dom.sortSelect.addEventListener('change', (e) => {
    state.sortKey = e.target.value;
    state.currentPage = 1;
    renderClientsTable();
  });

  // Smart Action Chips
  dom.smartChips.forEach((chip) => {
    chip.addEventListener('click', () => {
      dom.smartChips.forEach((c) => c.classList.remove('active'));
      chip.classList.add('active');
      state.smartFilter = chip.getAttribute('data-smart');
      state.currentPage = 1;
      renderClientsTable();
    });
  });

  // Pagination buttons
  dom.btnPrevPage.addEventListener('click', () => {
    if (state.currentPage > 1) {
      state.currentPage--;
      renderClientsTable();
    }
  });

  dom.btnNextPage.addEventListener('click', () => {
    state.currentPage++;
    renderClientsTable();
  });

  dom.btnResetFilters.addEventListener('click', () => {
    state.searchQuery = '';
    dom.clientSearchInput.value = '';
    dom.btnClearSearch.classList.add('hidden');
    state.segmentFilter = 'ALL';
    dom.segmentFilterPills.forEach((p) => {
      if (p.getAttribute('data-seg') === 'ALL') p.classList.add('active');
      else p.classList.remove('active');
    });
    state.smartFilter = 'ALL';
    dom.smartChips.forEach((c) => {
      if (c.getAttribute('data-smart') === 'ALL') c.classList.add('active');
      else c.classList.remove('active');
    });
    state.currentPage = 1;
    renderClientsTable();
  });

  // =========================================================================
  // CLIENT DRILL-DOWN MODAL & RECOMMENDATIONS
  // =========================================================================
  function openClientModal(clientCode) {
    const client = state.records.find((r) => r.code === clientCode);
    if (!client) return;

    state.selectedClientForModal = client;

    // Header info
    dom.modalClientName.textContent = client.nom;
    dom.modalClientCode.textContent = `#${client.code}`;
    dom.modalClientVendor.textContent = client.vendeur;
    dom.modalClientCA.textContent = fmt.currency(client.ca_ttc);

    dom.modalClientSegBadge.textContent = `SEG ${client.segment}`;
    dom.modalClientSegBadge.className = `badge-seg badge-seg-${client.segment.toLowerCase()}`;

    const target = client.msl_act + client.sku_rem;
    const activatedEntries = Object.entries(client.skus);
    const totalBoughtCount = activatedEntries.length;
    const mslAct = client.msl_act;
    const extraBought = Math.max(0, totalBoughtCount - mslAct);
    const totalMissingCount = state.allSkus.length - totalBoughtCount;
    const mslRemaining = client.sku_rem;
    const nonMslMissing = Math.max(0, totalMissingCount - mslRemaining);

    if (state.lang === 'ar') {
      dom.modalScoreText.textContent = `${mslAct} / ${target} (${client.taux_act.toFixed(1)}%)`;
      if (dom.modalCatalogBreakdownText) {
        dom.modalCatalogBreakdownText.innerHTML = `<strong>${mslAct}</strong> منتج إلزامي محقق من كوتا ${target} • إجمالي ما اشتراه المحل: <strong>${totalBoughtCount}</strong> منتجاً (${mslAct} كوتا + ${extraBought} إضافية)`;
      }

      dom.modalRemainingText.textContent = mslRemaining === 0 
        ? '✓ تم تحقيق 100% !' 
        : `${mslRemaining} منتج متبقي`;
      if (dom.modalRemainingSubText) {
        dom.modalRemainingSubText.innerHTML = mslRemaining === 0 
          ? '✓ المحل حقق 100% وهو مؤهل للاستفادة من خصم MSL الشهري (2%)' 
          : `طلب هذه الـ <strong>${mslRemaining}</strong> مواد إضافية يُمكّن المحل من الاستفادة من خصم 2% على كامل فاتورة الشهر`;
      }

      if (dom.modalActiveSkusSub) {
        dom.modalActiveSkusSub.textContent = `${mslAct} ضمن كوتا MSL • ${extraBought} مبيعات إضافية خارج الكوتا`;
      }
      if (dom.modalMissingSkusSub) {
        dom.modalMissingSkusSub.textContent = `منها ${mslRemaining} منتجات إلزامية لإنهاء كوتا 100% • ${nonMslMissing} مواد اختيارية`;
      }
    } else {
      dom.modalScoreText.textContent = `${mslAct} / ${target} Références MSL (${client.taux_act.toFixed(1)}%)`;
      if (dom.modalCatalogBreakdownText) {
        dom.modalCatalogBreakdownText.innerHTML = `<strong>${mslAct}</strong> validées sur quota ${target} • Total acheté : <strong>${totalBoughtCount}</strong> articles (${mslAct} quota + ${extraBought} hors quota)`;
      }

      dom.modalRemainingText.textContent = mslRemaining === 0 
        ? '✓ Quota 100% Atteint !' 
        : `${mslRemaining} Articles Restants`;
      if (dom.modalRemainingSubText) {
        dom.modalRemainingSubText.innerHTML = mslRemaining === 0 
          ? '✓ Quota 100% validé : le magasin débloque sa Remise MSL (2%)' 
          : `Plus que <strong>${mslRemaining}</strong> références pour faire gagner au magasin sa Remise de 2% sur le mois`;
      }

      if (dom.modalActiveSkusSub) {
        dom.modalActiveSkusSub.textContent = `${mslAct} dans le quota MSL • ${extraBought} hors quota`;
      }
      if (dom.modalMissingSkusSub) {
        dom.modalMissingSkusSub.textContent = `dont ${mslRemaining} obligatoires pour le 100% • ${nonMslMissing} optionnels`;
      }
    }

    dom.modalProgressBar.style.width = `${Math.min(100, client.taux_act)}%`;

    // 7 Essential Eligibility Pillars & MTO Evaluation
    const pillarsStatus = SEVEN_PILLARS.map((pillar) => {
      const matchedEntries = Object.entries(client.skus).filter(
        ([name, qty]) => pillar.matches(name) && Number(qty) > 0
      );
      const maxQty = matchedEntries.length > 0 ? Math.max(...matchedEntries.map(([_, q]) => Number(q))) : 0;
      const isMtoReached = maxQty >= pillar.mto;

      let status = 'missing';
      if (isMtoReached) {
        status = 'valid';
      } else if (matchedEntries.length > 0) {
        status = 'insufficient'; // e.g. bought 5L for HT when 6L is required!
      }

      return {
        ...pillar,
        status,
        isValid: isMtoReached,
        maxQty,
        matchedCount: matchedEntries.length,
      };
    });

    const validPillarsCount = pillarsStatus.filter((p) => p.isValid).length;
    const isAllPillarsValid = validPillarsCount === SEVEN_PILLARS.length;
    const nonValidPillars = pillarsStatus.filter((p) => !p.isValid);

    if (dom.modalPillarsStatusBadge) {
      if (isAllPillarsValid) {
        dom.modalPillarsStatusBadge.className = 'pillars-status-badge eligible';
        dom.modalPillarsStatusBadge.innerHTML = state.lang === 'ar'
          ? '✅ 7 / 7 أركان محققة وفق شرط MTO • المحل مؤهل لخصم 2% للبقال'
          : '✅ 7 / 7 Piliers MTO Validés • Magasin Éligible Remise 2%';
      } else {
        const detailsList = nonValidPillars
          .map((p) => {
            const pName = state.lang === 'ar' ? p.nameAr : p.nameFr;
            const mtoReq = state.lang === 'ar' ? p.mtoDisplayAr : p.mtoDisplayFr;
            if (p.status === 'insufficient') {
              return `${pName} (${fmt.tonnage(p.maxQty)} &lt; ${mtoReq})`;
            }
            return `${pName} (${state.lang === 'ar' ? 'غير مطلوب' : 'Non commandé'})`;
          })
          .join('، ');

        dom.modalPillarsStatusBadge.className = 'pillars-status-badge incomplete';
        dom.modalPillarsStatusBadge.innerHTML = state.lang === 'ar'
          ? `⚠️ ${validPillarsCount} / 7 محققة • ينقصه / غير كافٍ: <strong>${detailsList}</strong> لتأهيل الخصم`
          : `⚠️ ${validPillarsCount} / 7 Piliers MTO • Manque / Insuffisant : <strong>${detailsList}</strong>`;
      }
    }

    if (dom.modalPillarsGrid) {
      dom.modalPillarsGrid.innerHTML = pillarsStatus
        .map((p) => {
          const name = state.lang === 'ar' ? p.nameAr : p.nameFr;
          const mtoTag = state.lang === 'ar' ? p.mtoDisplayAr : p.mtoDisplayFr;
          const ruleText = state.lang === 'ar' ? p.ruleAr : p.ruleFr;

          let statusClass = 'missing';
          let checkHtml = '';

          if (p.status === 'valid') {
            statusClass = 'valid';
            checkHtml = `<span class="pillar-check" title="${ruleText}">✓ ${fmt.tonnage(p.maxQty)}</span>`;
          } else if (p.status === 'insufficient') {
            statusClass = 'insufficient';
            const tooltip = state.lang === 'ar'
              ? `الكمية غير كافية: طلب ${fmt.tonnage(p.maxQty)} فقط ولا تبلغ شرط ${mtoTag} الأدنى المطلوب (${ruleText})`
              : `Quantité insuffisante : ${fmt.tonnage(p.maxQty)} sur ${mtoTag} minimum requis (${ruleText})`;
            checkHtml = `<span class="pillar-check" title="${tooltip}">⚠️ ${fmt.tonnage(p.maxQty)} &lt; MTO</span>`;
          } else {
            statusClass = 'missing';
            const tooltip = state.lang === 'ar'
              ? `غير مطلوب إطلاقاً (شرط MTO: ${mtoTag})`
              : `Non commandé (MTO requis : ${mtoTag})`;
            checkHtml = `<span class="pillar-check" title="${tooltip}">✗ 0 (${state.lang === 'ar' ? 'ناقص' : 'Manque'})</span>`;
          }

          return `
            <div class="pillar-chip ${statusClass}" title="${name} — ${ruleText}">
              <div class="pillar-chip-main">
                <span class="pillar-chip-name">
                  <span>${p.icon}</span>
                  <span>${name}</span>
                </span>
                <span class="pillar-mto-tag">MTO: ${mtoTag}</span>
              </div>
              ${checkHtml}
            </div>
          `;
        })
        .join('');
    }

    // Activated SKUs (Display Tonnage / Volume in Kg)
    state.currentModalActivatedEntries = activatedEntries.sort((a, b) => b[1] - a[1]);
    state.modalActiveSearchQuery = '';
    state.modalMissingSearchQuery = '';
    if (dom.modalActiveSkusSearch) dom.modalActiveSkusSearch.value = '';
    if (dom.modalMissingSkusSearch) dom.modalMissingSkusSearch.value = '';

    // Recommendations Engine (Boost Missing / Insufficient Pillars to the Top)
    const segmentPeers = state.records.filter((r) => r.segment === client.segment && r.code !== client.code);
    const peerSkuPopularity = {};

    segmentPeers.forEach((peer) => {
      Object.keys(peer.skus).forEach((skuName) => {
        peerSkuPopularity[skuName] = (peerSkuPopularity[skuName] || 0) + 1;
      });
    });

    const missingSkus = state.allSkus
      .filter((skuName) => !client.skus[skuName])
      .map((skuName) => {
        const belongsToPillar = nonValidPillars.find((p) => p.matches(skuName));
        return {
          name: skuName,
          belongsToPillar,
          peerBuyerCount: peerSkuPopularity[skuName] || 0,
          peerPenetration: segmentPeers.length > 0 ? ((peerSkuPopularity[skuName] || 0) / segmentPeers.length) * 100 : 0,
        };
      })
      .sort((a, b) => {
        if (a.belongsToPillar && !b.belongsToPillar) return -1;
        if (!a.belongsToPillar && b.belongsToPillar) return 1;
        return b.peerBuyerCount - a.peerBuyerCount;
      });

    state.currentModalMissingSkus = missingSkus;

    // Render both lists (supports search & shows entire SKU name with no truncation)
    renderModalActiveList();
    renderModalMissingList();

    // Pitch bullets with sales closing argument for the vendor
    const top3Missing = missingSkus.slice(0, 3);
    if (top3Missing.length > 0) {
      dom.modalPitchRecommendations.innerHTML = top3Missing
        .map((sku, idx) => {
          const p = sku.belongsToPillar;
          const pillarTag = p
            ? (state.lang === 'ar'
                ? `<span class="pitch-tag" style="background:var(--warning-subtle); color:var(--warning); margin-right:6px;">⭐ ركن مشروط MTO (${p.nameAr} ≥ ${p.mtoDisplayAr})</span>`
                : `<span class="pitch-tag" style="background:var(--warning-subtle); color:var(--warning); margin-left:6px;">⭐ Pilier MTO (${p.nameFr} ≥ ${p.mtoDisplayFr})</span>`)
            : '';
          const pitchDesc = p
            ? (state.lang === 'ar'
                ? `طلب هذه المادة بشرط MTO (على الأقل <strong>${p.mtoDisplayAr}</strong>) يُفعّل ركن <strong>${p.nameAr}</strong> ويؤهل المحل لخصم الـ 2%!`
                : `Commander cet article au seuil MTO (au moins <strong>${p.mtoDisplayFr}</strong>) valide le pilier <strong>${p.nameFr}</strong> et rend le magasin éligible à la Remise 2% !`)
            : (state.lang === 'ar'
                ? `تطلبه <strong>${sku.peerPenetration.toFixed(0)}%</strong> من محلات هذا الصنف. خطوة إضافية نحو الـ 100% لخصم البقال.`
                : `Commandé par <strong>${sku.peerPenetration.toFixed(0)}%</strong> des confrères du segment (+1 SKU vers la Remise 2%).`);

          return `
            <div class="pitch-item">
              <span class="pitch-bullet">${idx + 1}.</span>
              <div>
                <div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap;">
                  <strong>${sku.name}</strong>
                  ${pillarTag}
                </div>
                <div style="font-size:0.78rem; color:var(--text-secondary); margin-top:2px;">
                  ${pitchDesc}
                </div>
              </div>
            </div>
          `;
        })
        .join('');
    } else {
      dom.modalPitchRecommendations.innerHTML = `
        <div class="text-success" style="font-size:0.86rem; font-weight:600;">
          ${state.lang === 'ar' ? 'تهانينا ! هذا المحل حقق جميع المنتجات المقترحة لصفه.' : 'Félicitations ! Ce client détient l\'ensemble des références de son segment.'}
        </div>
      `;
    }

    dom.clientModal.classList.remove('hidden');
  }

  function renderModalActiveList() {
    if (!state.currentModalActivatedEntries) return;
    const q = fmt.stripAccents((state.modalActiveSearchQuery || '').toLowerCase());
    let filtered = state.currentModalActivatedEntries;
    if (q) {
      filtered = filtered.filter(([name]) => fmt.stripAccents(name.toLowerCase()).includes(q));
    }
    if (dom.modalActiveSkusCount) {
      dom.modalActiveSkusCount.textContent = filtered.length;
    }
    if (filtered.length === 0) {
      const emptyMsg = state.modalActiveSearchQuery
        ? (i18n[state.lang]?.modalNoSkuFound || 'لم يتم العثور على أي مادة')
        : (state.lang === 'ar' ? 'لم يتم طلب أي منتج بعد.' : 'Aucun article commandé.');
      dom.modalActiveSkusList.innerHTML = `<p class="text-muted" style="font-size:0.8rem; padding:8px 4px;">${emptyMsg}</p>`;
    } else {
      dom.modalActiveSkusList.innerHTML = filtered
        .map(([name, amount]) => `
          <div class="sku-card-row">
            <span class="sku-card-name" title="${name}">${name}</span>
            <span class="sku-card-val">${fmt.tonnage(amount)}</span>
          </div>
        `)
        .join('');
    }
  }

  function renderModalMissingList() {
    if (!state.currentModalMissingSkus) return;
    const q = fmt.stripAccents((state.modalMissingSearchQuery || '').toLowerCase());
    let filtered = state.currentModalMissingSkus;
    if (q) {
      filtered = filtered.filter((item) => fmt.stripAccents(item.name.toLowerCase()).includes(q));
    }
    if (dom.modalMissingSkusCount) {
      dom.modalMissingSkusCount.textContent = filtered.length;
    }
    if (filtered.length === 0) {
      const emptyMsg = state.modalMissingSearchQuery
        ? (i18n[state.lang]?.modalNoSkuFound || 'لم يتم العثور على أي مادة')
        : (state.lang === 'ar' ? 'لا توجد منتجات ناقصة.' : 'Aucun article manquant.');
      dom.modalMissingSkusList.innerHTML = `<p class="text-muted" style="font-size:0.8rem; padding:8px 4px;">${emptyMsg}</p>`;
    } else {
      dom.modalMissingSkusList.innerHTML = filtered
        .map((item) => `
          <div class="sku-card-row">
            <span class="sku-card-name" title="${item.name}">${item.name}</span>
            <span class="sku-card-val text-warning" style="font-size:0.75rem">${item.peerPenetration.toFixed(0)}%</span>
          </div>
        `)
        .join('');
    }
  }

  function closeClientModal() {
    dom.clientModal.classList.add('hidden');
    state.selectedClientForModal = null;
  }

  dom.btnCloseModal.addEventListener('click', closeClientModal);
  dom.btnCloseModalBottom.addEventListener('click', closeClientModal);
  dom.clientModal.addEventListener('click', (e) => {
    if (e.target === dom.clientModal) closeClientModal();
  });

  if (dom.modalActiveSkusSearch) {
    dom.modalActiveSkusSearch.addEventListener('input', (e) => {
      state.modalActiveSearchQuery = e.target.value;
      renderModalActiveList();
    });
  }

  if (dom.modalMissingSkusSearch) {
    dom.modalMissingSkusSearch.addEventListener('input', (e) => {
      state.modalMissingSearchQuery = e.target.value;
      renderModalMissingList();
    });
  }

  dom.btnShareWhatsAppClient.addEventListener('click', () => {
    if (!state.selectedClientForModal) return;
    const c = state.selectedClientForModal;
    const target = c.msl_act + c.sku_rem;
    const text = state.lang === 'ar'
      ? `🏪 *بيانات المحل - ${c.nom}* (#${c.code})\n` +
        `🏷️ *الصنف :* ${c.segment} (الهدف ${target} منتج)\n` +
        `💰 *المبيعات :* ${fmt.currency(c.ca_ttc)}\n` +
        `🎯 *تفعيل MSL :* ${c.msl_act}/${target} (${c.taux_act.toFixed(1)}%)\n` +
        `📦 *المنتجات المتبقية :* ${c.sku_rem}\n`
      : `🏪 *FICHE CLIENT - ${c.nom}* (#${c.code})\n` +
        `🏷️ *Segment :* ${c.segment} (Objectif ${target} SKUs)\n` +
        `💰 *CA Réalisé :* ${fmt.currency(c.ca_ttc)}\n` +
        `🎯 *Activation MSL :* ${c.msl_act}/${target} (${c.taux_act.toFixed(1)}%)\n` +
        `📦 *Articles manquants :* ${c.sku_rem}\n`;

    navigator.clipboard.writeText(text).then(() => {
      showToast(state.lang === 'ar' ? 'تم نسخ بيانات المحل لواتساب !' : 'Fiche copiée pour WhatsApp !', 'success');
    });
  });

  // =========================================================================
  // TAB 3: PRODUCTS & SKU MATRIX
  // =========================================================================
  function getSkuCategory(skuName) {
    const s = skuName.toUpperCase();
    if (s.includes('HUILE') || s.includes('SOJA') || s.includes('FRITURE') || s.includes('ECOPRO') || ('PET' in s && (s.includes('1L') || s.includes('2L') || s.includes('5L')))) {
      return 'HUILES';
    }
    if (s.includes('DOUCHE') || s.includes('GEL') || s.includes('CREME') || s.includes('ALOE') || s.includes('ANTI BAC') || s.includes('LAIT') || s.includes('SAVON') || s.includes('120G X 4')) {
      return 'SOINS & HYGIENE';
    }
    if (s.includes('LESSIVE') || s.includes('VAISSELLE') || s.includes('PATE') || s.includes('MENJEL') || s.includes('230G')) {
      return 'ENTRETIEN';
    }
    if (s.includes('TARTINER') || s.includes('BARQUETTE') || s.includes('FORTIFIE') || s.includes('FEUILLETAGE') || s.includes('PATISSIERE') || s.includes('MAY') || s.includes('KETCH') || s.includes('MOUTARD')) {
      return 'ALIMENTAIRE';
    }
    return 'AUTRES';
  }

  function renderSkuMatrix() {
    const vendorRecords = getVendorRecords();
    const totalStores = vendorRecords.length;

    const skuAgg = {};
    state.allSkus.forEach((sku) => {
      skuAgg[sku] = { name: sku, buyers: 0, volume: 0, category: getSkuCategory(sku) };
    });

    vendorRecords.forEach((r) => {
      Object.entries(r.skus).forEach(([sku, amt]) => {
        if (!skuAgg[sku]) skuAgg[sku] = { name: sku, buyers: 0, volume: 0, category: getSkuCategory(sku) };
        skuAgg[sku].buyers++;
        skuAgg[sku].volume += amt;
      });
    });

    let skuList = Object.values(skuAgg);

    if (state.skuCategoryFilter !== 'ALL') {
      skuList = skuList.filter((s) => s.category === state.skuCategoryFilter);
    }

    if (state.skuSearchQuery) {
      const q = state.skuSearchQuery.toLowerCase().trim();
      skuList = skuList.filter((s) => s.name.toLowerCase().includes(q));
    }

    skuList.sort((a, b) => b.buyers - a.buyers);

    dom.skusTableBody.innerHTML = skuList
      .map((item) => {
        const penetration = totalStores > 0 ? (item.buyers / totalStores) * 100 : 0;
        let oppBadge = '';
        if (penetration >= 40) {
          oppBadge = `<span class="text-success font-mono font-bold" style="font-size:0.75rem">${state.lang === 'ar' ? '⭐ أساسي' : '⭐ Leader'}</span>`;
        } else if (penetration >= 10) {
          oppBadge = `<span class="text-warning font-mono font-bold" style="font-size:0.75rem">${state.lang === 'ar' ? '📈 في التطور' : '📈 Moyen'}</span>`;
        } else {
          oppBadge = `<span class="text-muted font-mono" style="font-size:0.75rem">${state.lang === 'ar' ? '⚠️ فرصة نمو' : '⚠️ Faible'}</span>`;
        }

        return `
        <tr>
          <td><strong style="color:var(--text-primary)">${item.name}</strong></td>
          <td><span class="badge-seg" style="font-size:0.7rem">${item.category}</span></td>
          <td class="text-center font-mono"><strong>${item.buyers}</strong> / ${totalStores}</td>
          <td>
            <div class="table-act-wrap">
              <span class="font-mono" style="font-size:0.76rem">${penetration.toFixed(1)}%</span>
              <div class="table-progress-track">
                <div class="table-progress-bar bar-blue" style="width:${Math.min(100, penetration)}%"></div>
              </div>
            </div>
          </td>
          <td class="text-right font-mono" style="font-weight:700">${fmt.tonnage(item.volume)}</td>
          <td class="text-center">${oppBadge}</td>
        </tr>
      `;
      })
      .join('');
  }

  dom.skuSearchInput.addEventListener('input', (e) => {
    state.skuSearchQuery = e.target.value;
    renderSkuMatrix();
  });

  dom.skuCategoryPills.forEach((pill) => {
    pill.addEventListener('click', () => {
      dom.skuCategoryPills.forEach((p) => p.classList.remove('active'));
      pill.classList.add('active');
      state.skuCategoryFilter = pill.getAttribute('data-cat');
      renderSkuMatrix();
    });
  });

  // =========================================================================
  // TAB 4: LEADERBOARD AGENCIA
  // =========================================================================
  function renderLeaderboard() {
    const vendorLeaderboard = state.vendors.map((v) => {
      const records = state.records.filter((r) => r.vendeur === v);
      const kpis = computeKPIs(records);
      return {
        name: v,
        ...kpis,
      };
    });

    vendorLeaderboard.sort((a, b) => b.totalCA - a.totalCA);

    dom.leaderboardGrid.innerHTML = vendorLeaderboard
      .map((v, idx) => {
        let rankClass = 'rank-default';
        if (idx === 0) rankClass = 'rank-gold';
        else if (idx === 1) rankClass = 'rank-silver';
        else if (idx === 2) rankClass = 'rank-bronze';

        const isCurrent = v.name === state.activeVendor;

        return `
        <div class="leaderboard-card ${isCurrent ? 'current-vendor' : ''}">
          <div class="leaderboard-top">
            <div style="display:flex; align-items:center; gap:10px">
              <div class="leaderboard-rank-badge ${rankClass}">#${idx + 1}</div>
              <div>
                <div class="leaderboard-vendor-name">${v.name}</div>
                <div style="font-size:0.75rem; color:var(--text-muted)">${v.totalClients} ${state.lang === 'ar' ? 'محل' : 'clients'}</div>
              </div>
            </div>
            ${isCurrent ? `<span class="text-success font-bold" style="font-size:0.75rem">${state.lang === 'ar' ? 'أنت' : 'Actif'}</span>` : ''}
          </div>

          <div class="leaderboard-stats-grid">
            <div>
              <div class="lead-stat-val text-primary">${fmt.currency(v.totalCA)}</div>
              <div class="lead-stat-lbl">${state.lang === 'ar' ? 'المبيعات' : 'CA TTC'}</div>
            </div>
            <div>
              <div class="lead-stat-val text-success">${fmt.pct(v.activationRate)}</div>
              <div class="lead-stat-lbl">${state.lang === 'ar' ? 'تفعيل MSL' : 'Taux MSL'}</div>
            </div>
            <div>
              <div class="lead-stat-val">${fmt.currency(v.avgBasket)}</div>
              <div class="lead-stat-lbl">${state.lang === 'ar' ? 'السلة' : 'Panier'}</div>
            </div>
          </div>

          <div style="display:flex; justify-content:space-between; align-items:center; font-size:0.78rem;">
            <span>⭐ <strong>${v.clients100}</strong> à 100%</span>
            <span>🚀 <strong>${v.quickWins}</strong> proches</span>
            <button class="btn btn-ghost btn-sm btn-select-leader" data-vendor="${v.name}">
              ${state.lang === 'ar' ? 'عرض المحلات ←' : 'Filtrer →'}
            </button>
          </div>
        </div>
      `;
      })
      .join('');

    dom.leaderboardGrid.querySelectorAll('.btn-select-leader').forEach((btn) => {
      btn.addEventListener('click', () => {
        switchVendor(btn.getAttribute('data-vendor'));
        switchTab('clients');
      });
    });
  }

  // =========================================================================
  // VIEW TABS SWITCHER
  // =========================================================================
  function switchTab(tabId) {
    state.activeTab = tabId;

    dom.tabBtns.forEach((btn) => {
      if (btn.getAttribute('data-tab') === tabId) btn.classList.add('active');
      else btn.classList.remove('active');
    });

    dom.tabPaneClients.classList.add('hidden');
    dom.tabPaneProducts.classList.add('hidden');
    dom.tabPaneLeaderboard.classList.add('hidden');

    if (tabId === 'clients') {
      dom.tabPaneClients.classList.remove('hidden');
      renderClientsTable();
    } else if (tabId === 'products') {
      dom.tabPaneProducts.classList.remove('hidden');
      renderSkuMatrix();
    } else if (tabId === 'leaderboard') {
      dom.tabPaneLeaderboard.classList.remove('hidden');
      renderLeaderboard();
    }
  }

  dom.tabBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      switchTab(btn.getAttribute('data-tab'));
    });
  });

  // =========================================================================
  // RENDER COMPLETE ACTIVE VENDOR VIEW
  // =========================================================================
  function renderCurrentVendorView() {
    const vendorRecords = getVendorRecords();
    const kpis = computeKPIs(vendorRecords);
    updateKPIDisplay(kpis);
    updateSmartFilterCounts(vendorRecords);

    if (state.activeTab === 'clients') renderClientsTable();
    else if (state.activeTab === 'products') renderSkuMatrix();
    else if (state.activeTab === 'leaderboard') renderLeaderboard();
  }

  // =========================================================================
  // EXPORT TO CLEAN MULTI-SHEET VENDOR EXCEL (.XLSX)
  // =========================================================================
  function exportFilteredVendorExcel() {
    const vendorRecords = getVendorRecords();
    if (vendorRecords.length === 0) {
      showToast(state.lang === 'ar' ? 'لا توجد بيانات للتصدير.' : 'Aucune donnée à exporter.', 'warning');
      return;
    }

    if (typeof XLSX === 'undefined') {
      showToast('Bibliothèque Excel indisponible.', 'danger');
      return;
    }

    // Sheet 1: Synthese_Vendeur
    const summaryRows = vendorRecords.map((r) => {
      const target = r.msl_act + r.sku_rem;
      return {
        'Code Client': r.code,
        'Nom Client': r.nom,
        Segment: r.segment,
        'CA TTC (MAD)': r.ca_ttc,
        'MSL Active': r.msl_act,
        'Taux Activation (%)': r.taux_act,
        'Objectif Segment': target,
        'SKUs Restants': r.sku_rem,
        'Articles Commandes (Kg)': Object.entries(r.skus)
          .map(([k, v]) => `${k} (${fmt.tonnage(v)})`)
          .join(', '),
      };
    });

    // Sheet 2: Matrice_Complete_MSL
    const matrixRows = vendorRecords.map((r) => {
      const row = {
        Agence: r.agence,
        Vendeur: r.vendeur,
        'Code Client': r.code,
        SEGMENT: r.segment,
        'Nombre MSL Activé': r.msl_act,
        "Taux d'activation": (r.taux_act / 100).toFixed(4),
        'Nbr Sku restant': r.sku_rem,
        'CA  TTC': r.ca_ttc,
        'Nom Client': r.nom,
      };
      state.allSkus.forEach((sku) => {
        row[sku] = r.skus[sku] !== undefined ? r.skus[sku] : '';
      });
      return row;
    });

    const workbook = XLSX.utils.book_new();

    const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
    XLSX.utils.book_append_sheet(workbook, wsSummary, 'Synthese Vendeur');

    const wsMatrix = XLSX.utils.json_to_sheet(matrixRows);
    XLSX.utils.book_append_sheet(workbook, wsMatrix, 'Matrice Detaillee MSL');

    const cleanVendorName = (state.activeVendor === 'ALL' ? 'Agence_Globale' : state.activeVendor)
      .replace(/\s+/g, '_')
      .replace(/[^a-zA-Z0-9_]/g, '');

    const outFileName = `Suivi_MSL_${cleanVendorName}.xlsx`;
    XLSX.writeFile(workbook, outFileName);
    showToast(state.lang === 'ar' ? `تم تصدير الملف بنجاح : ${outFileName}` : `Fichier exporté avec 2 feuilles : ${outFileName}`, 'success');
  }

  dom.btnExportVendorExcel.addEventListener('click', exportFilteredVendorExcel);

  dom.btnPrintRouteSheet.addEventListener('click', () => {
    window.print();
  });

  dom.btnCopyDailySummary.addEventListener('click', () => {
    const vendorRecords = getVendorRecords();
    const kpis = computeKPIs(vendorRecords);
    const vendorName = state.activeVendor === 'ALL' ? 'Agence' : state.activeVendor;

    const summary = state.lang === 'ar'
      ? `📊 *ملخص التتبع التجاري MSL*\n` +
        `👤 *البائع :* ${vendorName}\n` +
        `📍 *المحلات :* ${vendorRecords.length} محل (${kpis.activeClients} نشط)\n` +
        `💰 *المبيعات الكلية :* ${fmt.currency(kpis.totalCA)}\n` +
        `🎯 *نسبة MSL :* ${fmt.pct(kpis.activationRate)} (${fmt.int(kpis.mslActSum)}/${fmt.int(kpis.mslTargetSum)} منتج)\n` +
        `⭐ *محلات 100% :* ${kpis.clients100} محل\n` +
        `🚀 *فرص سريعة (≤3) :* ${kpis.quickWins} محل\n` +
        `\n_تطبيق Vento MSL_`
      : `📊 *POINT D'ACTIVITÉ SUIVI MSL*\n` +
        `👤 *Vendeur :* ${vendorName}\n` +
        `📍 *Portefeuille :* ${vendorRecords.length} clients (${kpis.activeClients} actifs)\n` +
        `💰 *Chiffre d'Affaires :* ${fmt.currency(kpis.totalCA)}\n` +
        `🎯 *Taux MSL :* ${fmt.pct(kpis.activationRate)} (${fmt.int(kpis.mslActSum)}/${fmt.int(kpis.mslTargetSum)} SKUs)\n` +
        `⭐ *100% Validés :* ${kpis.clients100} clients\n` +
        `🚀 *Opportunités Rapides (≤3 restants) :* ${kpis.quickWins} clients\n` +
        `\n_Généré via Vento MSL_`;

    navigator.clipboard.writeText(summary).then(() => {
      showToast(state.lang === 'ar' ? 'تم نسخ التقرير لواتساب !' : 'Résumé WhatsApp copié !', 'success');
    });
  });

  // =========================================================================
  // FILE UPLOAD & DRAG AND DROP
  // =========================================================================
  function handleFileInput(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (e) {
      parseExcelData(e.target.result, file.name);
    };
    reader.readAsArrayBuffer(file);
  }

  dom.excelFileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) handleFileInput(e.target.files[0]);
  });

  dom.excelFileInputDrop.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) handleFileInput(e.target.files[0]);
  });

  ['dragenter', 'dragover'].forEach((eventName) => {
    dom.dropZoneArea.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dom.dropZoneArea.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach((eventName) => {
    dom.dropZoneArea.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dom.dropZoneArea.classList.remove('dragover');
    });
  });

  dom.dropZoneArea.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    const file = dt.files[0];
    if (file) handleFileInput(file);
  });

  // =========================================================================
  // PWA SERVICE WORKER & PHONE INSTALLATION
  // =========================================================================
  let deferredInstallPrompt = null;

  function setupOfflinePwa() {
    // Register offline Service Worker
    if ('serviceWorker' in navigator) {
      window.addEventListener('load', () => {
        navigator.serviceWorker.register('./sw.js').catch((err) => {
          console.warn('ServiceWorker registration skipped:', err);
        });
      });
    }

    const isStandalone =
      window.matchMedia('(display-mode: standalone)').matches ||
      window.navigator.standalone === true;

    if (isStandalone) {
      if (dom.btnInstallPwa) dom.btnInstallPwa.classList.add('hidden');
      return;
    }

    // Android / Desktop Chrome native install prompt
    window.addEventListener('beforeinstallprompt', (e) => {
      e.preventDefault();
      deferredInstallPrompt = e;
      if (dom.btnInstallPwa) {
        dom.btnInstallPwa.classList.remove('hidden');
      }
    });

    // iOS Safari detection (shows install helper since iOS uses Share -> Add to Home Screen)
    const isIos = /iphone|ipad|ipod/i.test(window.navigator.userAgent);
    if (isIos && dom.btnInstallPwa) {
      dom.btnInstallPwa.classList.remove('hidden');
    }

    if (dom.btnInstallPwa) {
      dom.btnInstallPwa.addEventListener('click', async () => {
        if (deferredInstallPrompt) {
          deferredInstallPrompt.prompt();
          const { outcome } = await deferredInstallPrompt.userChoice;
          if (outcome === 'accepted') {
            dom.btnInstallPwa.classList.add('hidden');
          }
          deferredInstallPrompt = null;
        } else if (isIos) {
          showToast(
            state.lang === 'ar'
              ? 'لتثبيت التطبيق على آيفون: اضغط على زر المشاركة ⬆️ ثم اختر "إضافة إلى الشاشة الرئيسية"'
              : 'Sur iPhone : appuyez sur Partager ⬆️ puis sur "Sur l\'écran d\'accueil"',
            'info',
            6000
          );
        }
      });
    }

    window.addEventListener('appinstalled', () => {
      deferredInstallPrompt = null;
      if (dom.btnInstallPwa) dom.btnInstallPwa.classList.add('hidden');
      showToast(
        state.lang === 'ar'
          ? 'تم تثبيت التطبيق بنجاح! يعمل الآن بدون إنترنت.'
          : 'Application installée ! Fonctionne désormais sans internet.',
        'success',
        4000
      );
    });
  }

  // =========================================================================
  // INITIALIZATION
  // =========================================================================
  function init() {
    applyTheme(state.theme);
    applyLanguage(state.lang);
    setupInfoPopovers();
    setupOfflinePwa();

    // Restore dataset permanently saved on the phone (IndexedDB / localStorage)
    loadPersistentDataset().then((data) => {
      if (data && data.records && data.records.length > 0) {
        loadDatasetIntoApp(data.records, data.allSkus, data.fileName);
      }
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
