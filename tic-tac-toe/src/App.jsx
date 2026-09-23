import { useState } from 'react';

function Square({ value, isWinning, isPlayable, onSquareClick, label }) {
  const classes = [
    'square',
    value === 'X' ? 'is-x' : '',
    value === 'O' ? 'is-o' : '',
    isWinning ? 'is-winning' : '',
    isPlayable ? 'is-playable' : '',
  ]
    .filter(Boolean)
    .join(' ');

  return (
    <button
      type="button"
      className={classes}
      onClick={onSquareClick}
      aria-label={label}
    >
      {value}
    </button>
  );
}

function Board({ squares, winningLine, gameOver, onSquareClick }) {
  return (
    <div className="board" role="grid" aria-label="Tic-tac-toe board">
      {squares.map((value, index) => (
        <Square
          key={index}
          value={value}
          isWinning={winningLine.includes(index)}
          isPlayable={!gameOver && !value}
          onSquareClick={() => onSquareClick(index)}
          label={value ? `${value} on square ${index + 1}` : `Square ${index + 1}`}
        />
      ))}
    </div>
  );
}

function moveLocation(history, move) {
  if (move === 0) {
    return null;
  }

  const previous = history[move - 1];
  const current = history[move];
  const index = current.findIndex((value, i) => value !== previous[i]);

  if (index < 0) {
    return null;
  }

  return {
    player: current[index],
    row: Math.floor(index / 3) + 1,
    col: (index % 3) + 1,
  };
}

export default function Game() {
  const [history, setHistory] = useState([Array(9).fill(null)]);
  const [currentMove, setCurrentMove] = useState(0);
  const xIsNext = currentMove % 2 === 0;
  const currentSquares = history[currentMove];
  const result = calculateWinner(currentSquares);
  const winner = result?.winner ?? null;
  const winningLine = result?.line ?? [];
  const isDraw = !winner && currentSquares.every(Boolean);
  const gameOver = Boolean(winner) || isDraw;
  const viewingPast = currentMove !== history.length - 1;

  function handlePlay(nextSquares) {
    const nextHistory = [...history.slice(0, currentMove + 1), nextSquares];
    setHistory(nextHistory);
    setCurrentMove(nextHistory.length - 1);
  }

  function handleSquareClick(index) {
    if (calculateWinner(currentSquares) || currentSquares[index]) {
      return;
    }

    const nextSquares = currentSquares.slice();
    nextSquares[index] = xIsNext ? 'X' : 'O';
    handlePlay(nextSquares);
  }

  function jumpTo(nextMove) {
    setCurrentMove(nextMove);
  }

  function resetGame() {
    setHistory([Array(9).fill(null)]);
    setCurrentMove(0);
  }

  let statusLabel = 'Next player';
  let statusValue = xIsNext ? 'X' : 'O';
  let statusState = 'play';

  if (winner) {
    statusLabel = 'Winner';
    statusValue = winner;
    statusState = 'win';
  } else if (isDraw) {
    statusLabel = 'Result';
    statusValue = 'Draw';
    statusState = 'draw';
  }

  const moves = history.map((_, move) => {
    const location = moveLocation(history, move);
    const isCurrent = move === currentMove;
    const title = move === 0 ? 'Game start' : `Move ${move}`;
    const detail = location ? `${location.player} at ${location.row}, ${location.col}` : 'Empty board';

    if (isCurrent) {
      return (
        <li key={move} className="history-item is-current">
          <div className="history-copy">
            <span className="history-title">{title}</span>
            <span className="history-detail">{detail}</span>
          </div>
          <span className="history-badge">Now</span>
        </li>
      );
    }

    return (
      <li key={move} className="history-item">
        <button type="button" className="history-button" onClick={() => jumpTo(move)}>
          <span className="history-copy">
            <span className="history-title">{title}</span>
            <span className="history-detail">{detail}</span>
          </span>
        </button>
      </li>
    );
  });

  return (
    <div className="page">
      <header className="intro">
        <p className="eyebrow">React tutorial</p>
        <h1>Tic-Tac-Toe</h1>
      </header>

      <main className="game">
        <section className="panel" aria-label="Game board">
          <div className={`status is-${statusState}`} aria-live="polite">
            <span className="status-label">{statusLabel}</span>
            <span className={`status-value${statusValue === 'X' || statusValue === 'O' ? ` is-${statusValue.toLowerCase()}` : ''}`}>
              {statusValue}
            </span>
          </div>
          {viewingPast && (
            <p className="status-note">Viewing a previous position</p>
          )}
          <Board
            squares={currentSquares}
            winningLine={winningLine}
            gameOver={gameOver}
            onSquareClick={handleSquareClick}
          />
        </section>

        <aside className="panel history" aria-label="Move history">
          <div className="history-header">
            <h2>History</h2>
            <button
              type="button"
              className="reset"
              onClick={resetGame}
              disabled={history.length === 1}
            >
              Reset
            </button>
          </div>
          <ol className="history-list">{moves}</ol>
        </aside>
      </main>
    </div>
  );
}

function calculateWinner(squares) {
  const lines = [
    [0, 1, 2],
    [3, 4, 5],
    [6, 7, 8],
    [0, 3, 6],
    [1, 4, 7],
    [2, 5, 8],
    [0, 4, 8],
    [2, 4, 6],
  ];

  for (let i = 0; i < lines.length; i++) {
    const [a, b, c] = lines[i];
    if (squares[a] && squares[a] === squares[b] && squares[a] === squares[c]) {
      return { winner: squares[a], line: [a, b, c] };
    }
  }

  return null;
}
