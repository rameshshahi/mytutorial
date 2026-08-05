import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  createId,
  createScoreboard,
  defaultCategories,
  defaultGames,
  defaultTeams,
  getQuestionsForCategory,
} from '../data/quizData';

const storageKeys = {
  teams: 'quiz-league-teams',
  games: 'quiz-league-games',
  categories: 'quiz-league-categories',
  scoreboard: 'quiz-league-scoreboard',
  uploadedQuestions: 'quiz-league-uploaded-questions',
};

const readStoredValue = (key, fallback) => {
  try {
    const storedValue = window.localStorage.getItem(key);
    return storedValue ? JSON.parse(storedValue) : fallback;
  } catch {
    return fallback;
  }
};

function QuizLeague() {
  const [teams, setTeams] = useState(() => readStoredValue(storageKeys.teams, defaultTeams));
  const [games, setGames] = useState(() => readStoredValue(storageKeys.games, defaultGames));
  const [categories, setCategories] = useState(() => readStoredValue(storageKeys.categories, defaultCategories));
  const [selectedTeamId, setSelectedTeamId] = useState(defaultTeams[0].id);
  const [selectedGameId, setSelectedGameId] = useState(defaultGames[0].id);
  const [selectedCategoryId, setSelectedCategoryId] = useState(defaultCategories[0].id);
  const [selectedQuestionIndex, setSelectedQuestionIndex] = useState(null);
  const [askedQuestionIndexes, setAskedQuestionIndexes] = useState([]);
  const [isAnswerVisible, setIsAnswerVisible] = useState(false);
  const [isQuestionDialogOpen, setIsQuestionDialogOpen] = useState(false);

  const [teamNameInput, setTeamNameInput] = useState('');
  const [gameNameInput, setGameNameInput] = useState('');
  const [categoryInput, setCategoryInput] = useState('');

  const [editingTeamId, setEditingTeamId] = useState(null);
  const [editingGameId, setEditingGameId] = useState(null);
  const [editingCategoryId, setEditingCategoryId] = useState(null);
  const [activeSection, setActiveSection] = useState('team');

  const normalizeStoredScoreboard = (storedValue) => {
  if (!storedValue || typeof storedValue !== 'object') {
    return createScoreboard(defaultGames, defaultTeams, defaultCategories);
  }

  const isLegacyScoreboard = Object.values(storedValue).every(
    (teamScores) =>
      teamScores &&
      typeof teamScores === 'object' &&
      !Array.isArray(teamScores) &&
      Object.values(teamScores).every((value) => typeof value === 'number')
  );

  if (!isLegacyScoreboard) {
    return storedValue;
  }

  const normalized = createScoreboard(defaultGames, defaultTeams, defaultCategories);
  const defaultGameId = defaultGames[0].id;

  Object.entries(storedValue).forEach(([teamId, teamScores]) => {
    Object.entries(teamScores).forEach(([categoryId, value]) => {
      if (normalized[defaultGameId]?.[teamId]?.hasOwnProperty(categoryId)) {
        normalized[defaultGameId][teamId][categoryId] = Number(value) || 0;
      }
    });
  });

  return normalized;
};

  const [scoreboard, setScoreboard] = useState(() => normalizeStoredScoreboard(readStoredValue(storageKeys.scoreboard, null)));
  const [uploadedQuestionsByKey, setUploadedQuestionsByKey] = useState(() => readStoredValue(storageKeys.uploadedQuestions, {}));
  const [scoreInput, setScoreInput] = useState('');
  const [lastResult, setLastResult] = useState('');
  const [uploadMessage, setUploadMessage] = useState('');

  useEffect(() => {
    window.localStorage.setItem(storageKeys.teams, JSON.stringify(teams));
  }, [teams]);

  useEffect(() => {
    window.localStorage.setItem(storageKeys.games, JSON.stringify(games));
  }, [games]);

  useEffect(() => {
    window.localStorage.setItem(storageKeys.categories, JSON.stringify(categories));
  }, [categories]);

  useEffect(() => {
    setScoreboard((prev) => {
      const nextScoreboard = createScoreboard(games, teams, categories);

      games.forEach((game) => {
        const previousGameScores = prev[game.id] || {};

        teams.forEach((team) => {
          const previousTeamScores = previousGameScores[team.id] || {};

          categories.forEach((category) => {
            nextScoreboard[game.id][team.id][category.id] = previousTeamScores[category.id] || 0;
          });
        });
      });

      return nextScoreboard;
    });
  }, [games, teams, categories]);

  useEffect(() => {
    window.localStorage.setItem(storageKeys.scoreboard, JSON.stringify(scoreboard));
  }, [scoreboard]);

  useEffect(() => {
    window.localStorage.setItem(storageKeys.uploadedQuestions, JSON.stringify(uploadedQuestionsByKey));
  }, [uploadedQuestionsByKey]);

  const closeQuestionDialog = () => {
    setIsQuestionDialogOpen(false);
    setSelectedQuestionIndex(null);
    setIsAnswerVisible(false);
  };

  useEffect(() => {
    if (!isQuestionDialogOpen) {
      return undefined;
    }

    const handleEscapeKey = (event) => {
      if (event.key === 'Escape') {
        closeQuestionDialog();
      }
    };

    window.addEventListener('keydown', handleEscapeKey);
    return () => {
      window.removeEventListener('keydown', handleEscapeKey);
    };
  }, [isQuestionDialogOpen]);

  const selectedTeam = useMemo(() => teams.find((team) => team.id === selectedTeamId) || teams[0], [teams, selectedTeamId]);
  const selectedGame = useMemo(() => games.find((game) => game.id === selectedGameId) || games[0], [games, selectedGameId]);
  const selectedCategory = useMemo(() => categories.find((category) => category.id === selectedCategoryId) || categories[0], [categories, selectedCategoryId]);
  const activeQuestionKey = `${selectedGameId}-${selectedCategoryId}`;

  const questions = useMemo(() => {
    const importedQuestions = uploadedQuestionsByKey[activeQuestionKey];
    if (Array.isArray(importedQuestions) && importedQuestions.length > 0) {
      return importedQuestions;
    }

    return getQuestionsForCategory(selectedCategory?.name);
  }, [activeQuestionKey, uploadedQuestionsByKey, selectedCategory]);

  const inferCategoryIdForFile = (fileName) => {
    const lowerFileName = fileName.toLowerCase();
    const matchedCategory = categories.find((category) =>
      lowerFileName.includes(category.name.toLowerCase())
    );

    return matchedCategory ? matchedCategory.id : selectedCategoryId;
  };

  const visibleQuestions = useMemo(() => {
    return questions.filter((_, index) => !askedQuestionIndexes.includes(index));
  }, [questions, askedQuestionIndexes]);

  const currentQuestion = selectedQuestionIndex === null ? null : questions[selectedQuestionIndex];

  const normalizeCellValue = (value) => {
    if (value === null || value === undefined) {
      return '';
    }

    return String(value).trim();
  };

  const getQuestionRowsFromWorkbook = (sheet) => {
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

    return rows
      .map((row) => {
        const rowLookup = Object.keys(row || {}).reduce((accumulator, key) => {
          accumulator[key.toLowerCase()] = row[key];
          return accumulator;
        }, {});

        const questionText = normalizeCellValue(
          rowLookup.questionname ??
            rowLookup.question ??
            rowLookup.q ??
            rowLookup.prompt ??
            rowLookup.ques ??
            rowLookup.question_text ??
            rowLookup.questiontext
        );
        const answerText = normalizeCellValue(
          rowLookup.answer ??
            rowLookup.ans ??
            rowLookup.solution ??
            rowLookup.correctanswer ??
            rowLookup.correct_answer ??
            rowLookup.answer_text
        );

        if (!questionText || !answerText) {
          return null;
        }

        const options = [
          rowLookup.option1 ?? rowLookup.option_1 ?? rowLookup.a ?? rowLookup.optiona,
          rowLookup.option2 ?? rowLookup.option_2 ?? rowLookup.b ?? rowLookup.optionb,
          rowLookup.option3 ?? rowLookup.option_3 ?? rowLookup.c ?? rowLookup.optionc,
          rowLookup.option4 ?? rowLookup.option_4 ?? rowLookup.d ?? rowLookup.optiond,
        ]
          .map((option) => normalizeCellValue(option))
          .filter(Boolean);

        return {
          question: questionText,
          answer: answerText,
          ...(options.length > 0 ? { options } : {}),
        };
      })
      .filter(Boolean);
  };

  const handleExcelUpload = async (event) => {
    const selectedFiles = Array.from(event.target.files || []);
    if (selectedFiles.length === 0) {
      return;
    }

    const nextUploadedQuestionsByKey = { ...uploadedQuestionsByKey };
    const results = [];
    let loadedFiles = 0;

    for (const targetFile of selectedFiles) {
      try {
        const workbook = XLSX.read(await targetFile.arrayBuffer(), { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const parsedQuestions = getQuestionRowsFromWorkbook(sheet);

        if (parsedQuestions.length === 0) {
          results.push(`No readable rows found in ${targetFile.name}.`);
          continue;
        }

        const categoryId = inferCategoryIdForFile(targetFile.name);
        const workbookKey = `${selectedGameId}-${categoryId}`;

        nextUploadedQuestionsByKey[workbookKey] = [
          ...(nextUploadedQuestionsByKey[workbookKey] || []),
          ...parsedQuestions,
        ];
        loadedFiles += 1;
        const category = categories.find((item) => item.id === categoryId);
        results.push(`Loaded ${parsedQuestions.length} questions from ${targetFile.name} for ${selectedGame?.name} / ${category?.name || selectedCategory?.name}.`);
      } catch {
        results.push(`Could not read ${targetFile.name} as a worksheet.`);
      }
    }

    setUploadedQuestionsByKey(nextUploadedQuestionsByKey);
    setSelectedQuestionIndex(null);
    setIsAnswerVisible(false);
    setLastResult('');
    setUploadMessage(`Imported ${loadedFiles} file(s) for ${selectedGame?.name}. ${results.join(' ')}`);
    event.target.value = '';
  };

  const clearScoresForSelectedGame = () => {
    if (!window.confirm(`Clear all scores for ${selectedGame?.name}? This cannot be undone.`)) {
      return;
    }

    setScoreboard((prev) => ({
      ...prev,
      [selectedGameId]: teams.reduce((teamScores, team) => {
        teamScores[team.id] = categories.reduce((categoryScores, category) => {
          categoryScores[category.id] = 0;
          return categoryScores;
        }, {});
        return teamScores;
      }, {}),
    }));

    setLastResult(`Cleared all scores for ${selectedGame?.name}.`);
  };

  const deleteQuestionsForSelectedGame = () => {
    if (!window.confirm(`Delete uploaded questions for ${selectedGame?.name}? This cannot be undone.`)) {
      return;
    }

    setUploadedQuestionsByKey((prev) => {
      const next = { ...prev };
      Object.keys(next).forEach((key) => {
        if (key.startsWith(`${selectedGameId}-`)) {
          delete next[key];
        }
      });
      return next;
    });

    setSelectedQuestionIndex(null);
    setIsAnswerVisible(false);
    setLastResult(`Deleted uploaded questions for ${selectedGame?.name}.`);
  };

  const deleteSelectedTeamName = () => {
    if (!selectedTeamId) {
      return;
    }

    const teamToDelete = teams.find((team) => team.id === selectedTeamId);
    if (!teamToDelete) {
      return;
    }

    if (!window.confirm(`Delete team ${teamToDelete.name}? This will remove the team and its scores.`)) {
      return;
    }

    deleteTeam(selectedTeamId);
    setLastResult(`Deleted team ${teamToDelete.name}.`);
  };

  const saveTeam = () => {
    const trimmedName = teamNameInput.trim();
    if (!trimmedName) {
      return;
    }

    if (editingTeamId) {
      const isDuplicate = teams.some((team) => team.id !== editingTeamId && team.name.toLowerCase() === trimmedName.toLowerCase());
      if (isDuplicate) {
        return;
      }

      setTeams((prev) => prev.map((team) => (team.id === editingTeamId ? { ...team, name: trimmedName } : team)));
      setSelectedTeamId(editingTeamId);
      setEditingTeamId(null);
    } else {
      const isDuplicate = teams.some((team) => team.name.toLowerCase() === trimmedName.toLowerCase());
      if (isDuplicate) {
        return;
      }

      const newTeam = { id: createId(), name: trimmedName };
      setTeams((prev) => [...prev, newTeam]);
      setSelectedTeamId(newTeam.id);
    }

    setTeamNameInput('');
  };

  const deleteTeam = (teamId) => {
    const nextTeams = teams.filter((team) => team.id !== teamId);
    setTeams(nextTeams);

    if (selectedTeamId === teamId) {
      setSelectedTeamId(nextTeams[0]?.id || '');
    }

    if (editingTeamId === teamId) {
      setEditingTeamId(null);
      setTeamNameInput('');
    }
  };

  const startEditingTeam = (team) => {
    setEditingTeamId(team.id);
    setTeamNameInput(team.name);
  };

  const saveGame = () => {
    const trimmedName = gameNameInput.trim();
    if (!trimmedName) {
      return;
    }

    if (editingGameId) {
      const isDuplicate = games.some((game) => game.id !== editingGameId && game.name.toLowerCase() === trimmedName.toLowerCase());
      if (isDuplicate) {
        return;
      }

      setGames((prev) => prev.map((game) => (game.id === editingGameId ? { ...game, name: trimmedName } : game)));
      setSelectedGameId(editingGameId);
      setEditingGameId(null);
    } else {
      const isDuplicate = games.some((game) => game.name.toLowerCase() === trimmedName.toLowerCase());
      if (isDuplicate) {
        return;
      }

      const newGame = { id: createId(), name: trimmedName };
      setGames((prev) => [...prev, newGame]);
      setSelectedGameId(newGame.id);
    }

    setGameNameInput('');
  };

  const deleteGame = (gameId) => {
    const nextGames = games.filter((game) => game.id !== gameId);
    setGames(nextGames);

    if (selectedGameId === gameId) {
      setSelectedGameId(nextGames[0]?.id || '');
    }

    if (editingGameId === gameId) {
      setEditingGameId(null);
      setGameNameInput('');
    }
  };

  const startEditingGame = (game) => {
    setEditingGameId(game.id);
    setGameNameInput(game.name);
  };

  const saveCategory = () => {
    const trimmedName = categoryInput.trim();
    if (!trimmedName) {
      return;
    }

    if (editingCategoryId) {
      const isDuplicate = categories.some((category) => category.id !== editingCategoryId && category.name.toLowerCase() === trimmedName.toLowerCase());
      if (isDuplicate) {
        return;
      }

      setCategories((prev) => prev.map((category) => (category.id === editingCategoryId ? { ...category, name: trimmedName } : category)));
      setSelectedCategoryId(editingCategoryId);
      setEditingCategoryId(null);
    } else {
      const isDuplicate = categories.some((category) => category.name.toLowerCase() === trimmedName.toLowerCase());
      if (isDuplicate) {
        return;
      }

      const newCategory = { id: createId(), name: trimmedName };
      setCategories((prev) => [...prev, newCategory]);
      setSelectedCategoryId(newCategory.id);
    }

    setCategoryInput('');
    setSelectedQuestionIndex(0);
    setIsAnswerVisible(false);
    setLastResult('');
  };

  const deleteCategory = (categoryId) => {
    const nextCategories = categories.filter((category) => category.id !== categoryId);
    setCategories(nextCategories);

    if (selectedCategoryId === categoryId) {
      setSelectedCategoryId(nextCategories[0]?.id || '');
    }

    if (editingCategoryId === categoryId) {
      setEditingCategoryId(null);
      setCategoryInput('');
    }
  };

  const startEditingCategory = (category) => {
    setEditingCategoryId(category.id);
    setCategoryInput(category.name);
  };

  const saveScore = () => {
    const numericScore = Number(scoreInput);

    if (!Number.isFinite(numericScore) || numericScore < 0) {
      return;
    }

    const previousScore = Number(scoreboard?.[selectedGameId]?.[selectedTeam.id]?.[selectedCategory.id] || 0);
    const nextScore = previousScore + numericScore;

    setScoreboard((prev) => ({
      ...prev,
      [selectedGameId]: {
        ...prev[selectedGameId],
        [selectedTeam.id]: {
          ...prev[selectedGameId]?.[selectedTeam.id],
          [selectedCategory.id]: nextScore,
        },
      },
    }));

    setScoreInput('');
    setLastResult('Score added for ' + selectedTeam.name + ' in ' + selectedCategory.name + '. New total: ' + nextScore + '.');
  };

  return (
    <div className="app-shell">
      <div className="app-card">
        <h1>Quiz League</h1>
        <p className="subtitle">Each management area supports complete CRUD actions for teams, games, and categories.</p>

        <nav className="section-nav" aria-label="Quiz sections">
          {[
            { key: 'team', label: 'Team Management' },
            { key: 'game', label: 'Game Management' },
            { key: 'category', label: 'Category Management' },
            { key: 'play', label: 'Play Game' },
            { key: 'load', label: 'Load Question' },
            { key: 'clear', label: 'Clear Data' },
            { key: 'score', label: 'Score' },
          ].map((section) => (
            <button
              key={section.key}
              type="button"
              className={`nav-tab ${activeSection === section.key ? 'active' : ''}`}
              onClick={() => setActiveSection(section.key)}
            >
              {section.label}
            </button>
          ))}
        </nav>

        {activeSection === 'team' && (
          <div className="entry-card section-card">
            <h3>Team Management</h3>
            <label>
              Team name
              <div className="input-row">
                <input
                  type="text"
                  value={teamNameInput}
                  onChange={(e) => setTeamNameInput(e.target.value)}
                  placeholder="Enter team name"
                />
                <button type="button" onClick={saveTeam}>{editingTeamId ? 'Update team' : 'Add team'}</button>
              </div>
            </label>

            {editingTeamId && (
              <button
                type="button"
                className="secondary-btn"
                onClick={() => {
                  setEditingTeamId(null);
                  setTeamNameInput('');
                }}
              >
                Cancel edit
              </button>
            )}

            <label>
              Select team
              <select value={selectedTeamId} onChange={(e) => setSelectedTeamId(e.target.value)}>
                {teams.map((team) => (
                  <option key={team.id} value={team.id}>
                    {team.name}
                  </option>
                ))}
              </select>
            </label>

            <ul className="management-list">
              {teams.map((team) => (
                <li key={team.id}>
                  <span>{team.name}</span>
                  <div className="row-actions">
                    <button type="button" className="ghost-btn" onClick={() => startEditingTeam(team)}>Edit</button>
                    <button type="button" className="danger-btn" onClick={() => deleteTeam(team.id)}>Delete</button>
                  </div>
                </li>
              ))}
            </ul>
          </div>
        )}

        {activeSection === 'game' && (
          <div className="entry-card section-card">
            <h3>Game Management</h3>
            <label>
              Game name
              <div className="input-row">
                <input
                  type="text"
                  value={gameNameInput}
                  onChange={(e) => setGameNameInput(e.target.value)}
                  placeholder="Enter game name"
                />
                <button type="button" onClick={saveGame}>{editingGameId ? 'Update game' : 'Add game'}</button>
              </div>
            </label>

            {editingGameId && (
              <button
                type="button"
                className="secondary-btn"
                onClick={() => {
                  setEditingGameId(null);
                  setGameNameInput('');
                }}
              >
                Cancel edit
              </button>
            )}

            <label>
              Select game
              <select value={selectedGameId} onChange={(e) => setSelectedGameId(e.target.value)}>
                {games.map((game) => (
                  <option key={game.id} value={game.id}>
                    {game.name}
                  </option>
                ))}
              </select>
            </label>

            <ul className="management-list">
              {games.map((game) => (
                <li key={game.id}>
                  <span>{game.name}</span>
                  <div className="row-actions">
                    <button type="button" className="ghost-btn" onClick={() => startEditingGame(game)}>Edit</button>
                    <button type="button" className="danger-btn" onClick={() => deleteGame(game.id)}>Delete</button>
                  </div>
                </li>
              ))}
            </ul>
          </div>
        )}

        {activeSection === 'category' && (
          <div className="entry-card section-card">
            <h3>Category Management</h3>
            <label>
              Category name
              <div className="input-row">
                <input
                  type="text"
                  value={categoryInput}
                  onChange={(e) => setCategoryInput(e.target.value)}
                  placeholder="Enter category"
                />
                <button type="button" onClick={saveCategory}>{editingCategoryId ? 'Update category' : 'Add category'}</button>
              </div>
            </label>

            {editingCategoryId && (
              <button
                type="button"
                className="secondary-btn"
                onClick={() => {
                  setEditingCategoryId(null);
                  setCategoryInput('');
                }}
              >
                Cancel edit
              </button>
            )}

            <label>
              Select category
              <select
                value={selectedCategoryId}
                onChange={(e) => {
                  setSelectedCategoryId(e.target.value);
                  setSelectedQuestionIndex(0);
                  setIsAnswerVisible(false);
                  setLastResult('');
                }}
              >
                {categories.map((category) => (
                  <option key={category.id} value={category.id}>
                    {category.name}
                  </option>
                ))}
              </select>
            </label>

            <ul className="management-list">
              {categories.map((category) => (
                <li key={category.id}>
                  <span>{category.name}</span>
                  <div className="row-actions">
                    <button type="button" className="ghost-btn" onClick={() => startEditingCategory(category)}>Edit</button>
                    <button type="button" className="danger-btn" onClick={() => deleteCategory(category.id)}>Delete</button>
                  </div>
                </li>
              ))}
            </ul>
          </div>
        )}

        {activeSection === 'play' && (
          <>
            <div className="entry-card section-card">
              <h3>Play Game</h3>
              <div className="play-controls">
                <label>
                  Select game
                  <select value={selectedGameId} onChange={(e) => setSelectedGameId(e.target.value)}>
                    {games.map((game) => (
                      <option key={game.id} value={game.id}>
                        {game.name}
                      </option>
                    ))}
                  </select>
                </label>

                <label>
                  Select category
                  <select
                    value={selectedCategoryId}
                    onChange={(e) => {
                      setSelectedCategoryId(e.target.value);
                      setSelectedQuestionIndex(null);
                      setIsAnswerVisible(false);
                      setLastResult('');
                    }}
                  >
                    {categories.map((category) => (
                      <option key={category.id} value={category.id}>
                        {category.name}
                      </option>
                    ))}
                  </select>
                </label>
              </div>
            </div>

            <div className="quiz-panel section-card">
              <div className="question-tile-grid">
                {visibleQuestions.map((question, visibleIndex) => {
                  const actualIndex = questions.findIndex((item) => item === question);
                  return (
                    <button
                      key={actualIndex}
                      type="button"
                      className={`question-tile ${selectedQuestionIndex === actualIndex ? 'selected' : ''}`}
                      onClick={() => {
                        setSelectedQuestionIndex(actualIndex);
                        setAskedQuestionIndexes((prev) => (prev.includes(actualIndex) ? prev : [...prev, actualIndex]));
                        setIsAnswerVisible(false);
                        setIsQuestionDialogOpen(true);
                      }}
                    >
                      Question {actualIndex + 1}
                    </button>
                  );
                })}
              </div>

              {isQuestionDialogOpen && currentQuestion && (
                <div className="question-dialog-overlay" onClick={closeQuestionDialog}>
                  <div className="question-dialog" onClick={(event) => event.stopPropagation()}>
                    <div className="question-dialog-header">
                      <p className="question-label">{selectedGame?.name} • {selectedCategory?.name} question {selectedQuestionIndex + 1}</p>
                      <button
                        type="button"
                        className="secondary-btn close-dialog-btn"
                        onClick={closeQuestionDialog}
                      >
                        Close
                      </button>
                    </div>

                    <h2>{currentQuestion.question}</h2>
                    <button
                      type="button"
                      className="check-answer-btn"
                      onClick={() => setIsAnswerVisible((prev) => !prev)}
                    >
                      {isAnswerVisible ? 'Hide Answer' : 'Check Answer'}
                    </button>

                    {isAnswerVisible && (
                      <div className="answer-box">
                        <strong>Answer:</strong> {currentQuestion.answer}
                      </div>
                    )}
                  </div>
                </div>
              )}

              {lastResult && <p className="result-text">{lastResult}</p>}
            </div>
          </>
        )}

        {activeSection === 'load' && (
          <>
            <div className="entry-card section-card">
              <h3>Load Question</h3>
              <div className="play-controls">
                <label>
                  Select game
                  <select value={selectedGameId} onChange={(e) => setSelectedGameId(e.target.value)}>
                    {games.map((game) => (
                      <option key={game.id} value={game.id}>
                        {game.name}
                      </option>
                    ))}
                  </select>
                </label>

                <label>
                  Select category
                  <select
                    value={selectedCategoryId}
                    onChange={(e) => {
                      setSelectedCategoryId(e.target.value);
                      setSelectedQuestionIndex(null);
                      setIsAnswerVisible(false);
                      setLastResult('');
                    }}
                  >
                    {categories.map((category) => (
                      <option key={category.id} value={category.id}>
                        {category.name}
                      </option>
                    ))}
                  </select>
                </label>
              </div>

              <div className="input-row upload-row">
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  multiple
                  onChange={handleExcelUpload}
                />
              </div>

              <p className="result-text">
                {uploadMessage || `Upload an Excel workbook for ${selectedGame?.name} / ${selectedCategory?.name} and it will be loaded for that specific selection.`}
              </p>

              {uploadedQuestionsByKey[activeQuestionKey] && (
                <p className="result-text">
                  Loaded {uploadedQuestionsByKey[activeQuestionKey].length} question(s) for the current game/category pair.
                </p>
              )}

              <div className="input-row" style={{ marginTop: '16px' }}>
                <button type="button" className="danger-btn" onClick={deleteQuestionsForSelectedGame}>
                  Clear all uploaded questions for selected game
                </button>
              </div>
            </div>
          </>
        )}

        {activeSection === 'clear' && (
          <div className="entry-card section-card">
            <h3>Clear Data</h3>
            <div className="play-controls">
              <label>
                Select game
                <select value={selectedGameId} onChange={(e) => setSelectedGameId(e.target.value)}>
                  {games.map((game) => (
                    <option key={game.id} value={game.id}>
                      {game.name}
                    </option>
                  ))}
                </select>
              </label>

              <label>
                Select team
                <select value={selectedTeamId} onChange={(e) => setSelectedTeamId(e.target.value)}>
                  {teams.map((team) => (
                    <option key={team.id} value={team.id}>
                      {team.name}
                    </option>
                  ))}
                </select>
              </label>
            </div>

            <div className="input-row">
              <button type="button" className="danger-btn" onClick={clearScoresForSelectedGame}>
                Clear all scores for selected game
              </button>
              <button type="button" className="danger-btn" onClick={deleteQuestionsForSelectedGame}>
                Delete uploaded questions for selected game
              </button>
            </div>

            <div className="input-row" style={{ marginTop: '16px' }}>
              <button type="button" className="danger-btn" onClick={deleteSelectedTeamName}>
                Delete selected team
              </button>
            </div>

            {lastResult && <p className="result-text">{lastResult}</p>}
          </div>
        )}

        {activeSection === 'score' && (
          <div className="scoreboard-card section-card">
            <h2>Score Entry</h2>
            <div className="score-entry-grid">
              <label>
                Select game
                <select value={selectedGameId} onChange={(e) => setSelectedGameId(e.target.value)}>
                  {games.map((game) => (
                    <option key={game.id} value={game.id}>
                      {game.name}
                    </option>
                  ))}
                </select>
              </label>

              <label>
                Select category
                <select value={selectedCategoryId} onChange={(e) => setSelectedCategoryId(e.target.value)}>
                  {categories.map((category) => (
                    <option key={category.id} value={category.id}>
                      {category.name}
                    </option>
                  ))}
                </select>
              </label>

              <label>
                Select team
                <select value={selectedTeamId} onChange={(e) => setSelectedTeamId(e.target.value)}>
                  {teams.map((team) => (
                    <option key={team.id} value={team.id}>
                      {team.name}
                    </option>
                  ))}
                </select>
              </label>

              <label>
                Score
                <div className="input-row">
                  <input
                    type="number"
                    min="0"
                    value={scoreInput}
                    onChange={(e) => setScoreInput(e.target.value)}
                    placeholder="Enter score"
                  />
                  <button type="button" onClick={saveScore}>Save Score</button>
                </div>
              </label>
            </div>

            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Team</th>
                    {categories.map((category) => (
                      <th key={category.id}>{category.name}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {teams.map((team) => (
                    <tr key={team.id}>
                      <td>{team.name}</td>
                      {categories.map((category) => (
                        <td key={category.id}>{scoreboard?.[selectedGameId]?.[team.id]?.[category.id] || 0}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {lastResult && <p className="result-text">{lastResult}</p>}
          </div>
        )}
      </div>
    </div>
  );
}

export default QuizLeague;
