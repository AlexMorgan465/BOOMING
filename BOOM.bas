/* ---------------- DATA : Mention Bac ---------------- */

const mb_managers = [
  {
    zone:'ZONE6',
    role:'Back-UP Manager',
    name:'RHOUDANE OTHMANE',
    days:[
      w('8:00','18:00'),
      w('8:00','18:00'),
      w('8:00','18:00'),
      w('8:00','18:00'),
      w('8:00','17:00'),
      OFF,
      OFF
    ],
    heures:'44:00:00',
    comment:'RAS'
  },
  {
    zone:'ZONE4',
    role:'Missionnée',
    name:'IHASSANE IKRAM',
    days:[
      OFF,
      w('9:00','19:00'),
      w('8:00','18:00'),
      w('9:00','19:00'),
      w('10:00','19:00'),
      w('9:00','19:00'),
      OFF
    ],
    heures:'44:00:00',
    comment:'RAS'
  }
];

const mb_pause = [];

const mb_zones = [
  // <<< Paste all collaborators from your Excel here >>>
];

const mb_stats = {
  ouverture:[24,22,21,25,16,7,0],
  renfort:[0,0,0,0,0,0,0],
  middle:[0,0,0,0,4,7,0],
  middlePlus:[0,0,0,0,0,0,0],
  fermeture:[2,4,5,1,2,12,26]
};
