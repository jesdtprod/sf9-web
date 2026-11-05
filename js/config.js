/* =====================================================================
   E-CARD settings — edit this file, no other code needs to change.
   ===================================================================== */
window.ECARD_CONFIG = {
  school: {
    name: 'Notre Dame of Sarmiento, Inc.',
    address: 'Sarmiento, Parang, Maguindanao Del Norte',
    title: "LEARNER'S REPORT CARD",
    subtitle: 'Senior High School',
    principal: '',
    adviser: ''
  },
  templates: {
    ndp: { name: 'Notre Dame of Parang, Inc.', address: 'Mabini St., Parang, Maguindanao Del Norte', recognition: 'Government Recognition No. 139 s., 1967', logo: 'assets/ndp.png' },
    nds: { name: 'Notre Dame of Sarmiento, Inc.', address: 'Sarmiento, Parang, Maguindanao Del Norte', recognition: 'Government Recognition No. 40 s., 1971', logo: 'assets/nds.png' }
  },

  // sheet names inside the Excel file
  sheets: { students: 'Students', values: 'Core Values', sem1: '1st Sem', sem2: '2nd Sem', attendance: 'Attendance' },

  passMark: 75,                 // PASSED when grade > passMark - 0.1  (same as your > 74.9)
  requireBothQuarters: true,    // final grade only shows when both quarters are filled in
  excludeFromAverage: ['Homeroom', 'Conduct'],   // same as your General Average formula

  // how the subjects of each semester sheet are grouped on the card (in sheet order)
  layout: {
    1: [['Core Subjects', 5], ['Applied / Contextualized Subject', 1], ['Specialized Subjects', 2], ['Other Subjects', 3]],
    2: [['Core Subjects', 3], ['Applied / Contextualized Subject', 2], ['Specialized Subjects', 3], ['Other Subjects', 3]]
  },

  // one entry per block in the "Core Values" sheet, left to right
  values: [
    { code: 'N', name: 'Noble', statement: 'A Notre Damean demonstrates high moral principles that embody nobility of character especially honesty and integrity.' },
    { code: 'O', name: 'Obedient', statement: 'A Notre Damean exemplifies fidelity to God and loyalty to country; follows school policies, rules and regulations; and practices desirable family values.' },
    { code: 'T', name: 'Transformed', statement: 'A Notre Damean displays positive disposition in life and an acceptable level of self-confidence and maturity, and shows capacity for teamwork, collaboration and solidarity.' },
    { code: 'R', name: 'Respectful', statement: 'A Notre Damean shows high esteem and regard to oneself, others, the community, country, nature, environment, and others faiths and cultures.' },
    { code: 'E', name: 'Enlightened', statement: "A Notre Damean demonstrates academic excellence, critical and reflective thinking, and guided by one's conscience, makes good decisions and actions based on wise judgment." },
    { code: 'D', name: 'Disciplined', statement: 'A Notre Damean wills and obliges to do what is right and expected, and exercises self-control in one’s dealings with the world and others.' },
    { code: 'A', name: 'Accountable', statement: 'A Notre Damean displays commitment and responsibility in performing one’s tasks, and shows leadership capacities and sense of volunteerism and initiative.' },
    { code: 'M', name: 'Marian', statement: 'A Notre Damean emulates Mary’s virtues especially faith in God, generosity, humility, prayerfulness, purity, simplicity, and service.' },
    { code: 'E', name: 'Empathetic', statement: 'A Notre Damean shows sensitivity and understanding of the feelings, ideas, beliefs, cultures, traditions, and experience of others; and shows compassion, friendship, kindness and love.' }
  ],

  markings: [['AO', 'Always Observed'], ['SO', 'Sometimes Observed'], ['RO', 'Rarely Observed'], ['NO', 'Not Observed']],
  scale: [
    ['Outstanding', '90-100', 'Passed'], ['Very Satisfactory', '85-89', 'Passed'], ['Satisfactory', '80-84', 'Passed'],
    ['Fairly Satisfactory', '75-79', 'Passed'], ['Did not Meet Expectations', 'Below 75', 'Failed']
  ]
};
