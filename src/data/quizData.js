export const createId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);

export const defaultTeams = [
  { id: 'team-1', name: 'Team A' },
  { id: 'team-2', name: 'Team B' },
  { id: 'team-3', name: 'Team C' },
  { id: 'team-4', name: 'Team D' },
];

export const defaultCategories = [
  { id: 'cat-1', name: 'Sport' },
  { id: 'cat-2', name: 'History' },
];

export const defaultGames = [{ id: 'game-1', name: 'Main Game' }];

export const quizQuestions = {
  Sport: [
    {
      question: 'Which sport uses the terms love and deuce?',
      options: ['Cricket', 'Tennis', 'Basketball', 'Baseball'],
      answer: 'Tennis',
    },
    {
      question: 'How many players are on a standard soccer team on the field?',
      options: ['9', '10', '11', '12'],
      answer: '11',
    },
  ],
  History: [
    {
      question: 'Who was the first President of the United States?',
      options: ['Abraham Lincoln', 'George Washington', 'Thomas Jefferson', 'John Adams'],
      answer: 'George Washington',
    },
    {
      question: 'Which ancient civilization built the pyramids?',
      options: ['Roman', 'Greek', 'Egyptian', 'Viking'],
      answer: 'Egyptian',
    },
  ],
};

export const createScoreboard = (gameList, teamList, categoryList) => {
  return gameList.reduce((scoreboard, game) => {
    scoreboard[game.id] = teamList.reduce((teamScores, team) => {
      teamScores[team.id] = categoryList.reduce((categoryScores, category) => {
        categoryScores[category.id] = 0;
        return categoryScores;
      }, {});
      return teamScores;
    }, {});
    return scoreboard;
  }, {});
};

export const getQuestionsForCategory = (categoryName) => {
  if (quizQuestions[categoryName]) {
    return quizQuestions[categoryName];
  }

  return [
    {
      question: 'Sample question for ' + categoryName + ' category?',
      options: ['Option 1', 'Option 2', 'Option 3', 'Option 4'],
      answer: 'Option 1',
    },
  ];
};
