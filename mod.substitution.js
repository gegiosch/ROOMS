var Mod = Mod || {};

Mod.Substitution = {
  handlePage: function () {
    var actor = ROOMS_APP.Auth.getUserContext();
    var bootstrap = {
      page: 'substitution',
      title: 'Sostituzioni',
      subtitle: 'Modulo unificato per la gestione operativa e la consultazione dei report.',
      user: actor,
      tabs: [
        {
          id: 'daily',
          label: 'Sostituzioni giornaliere',
          kind: 'iframe',
          src: '?page=board&substitutionStandalone=true&substitutionTab=daily'
        },
        {
          id: 'absences',
          label: 'Inserimento assenze',
          kind: 'placeholder'
        },
        {
          id: 'trips',
          label: 'Uscite didattiche',
          kind: 'iframe',
          src: '?page=board&substitutionStandalone=true&substitutionTab=trips'
        },
        {
          id: 'long',
          label: 'Supplenze lunghe',
          kind: 'iframe',
          src: '?page=board&substitutionStandalone=true&substitutionTab=long'
        },
        {
          id: 'reports',
          label: 'Report giornalieri',
          kind: 'iframe',
          src: '?fn=substitutionReports&embedded=true'
        }
      ]
    };

    return renderTemplate_('ui.substitution', {
      pageTitle: 'Sostituzioni',
      initialModelJson: JSON.stringify(bootstrap)
    });
  }
};
