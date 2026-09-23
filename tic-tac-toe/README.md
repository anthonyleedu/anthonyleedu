# Tic-Tac-Toe

A complete Tic-Tac-Toe game based on the [official React tutorial](https://react.dev/learn/tutorial-tic-tac-toe).

The app includes:

- A 3×3 board where X and O take turns
- Winner detection
- Move history (“time travel”) so you can jump back to any previous turn

The game logic matches the finished tutorial (`Square`, `Board`, `Game`, and `calculateWinner`). The project is packaged with [Vite](https://vite.dev/) so you can install dependencies and run it locally.

## Prerequisites

- [Node.js](https://nodejs.org/) 18 or later (tested with Node 22)
- npm (bundled with Node.js)

## Setup

From this `tic-tac-toe` directory:

```bash
npm install
```

## Run the development server

```bash
npm run dev
```

Then open [http://localhost:5173](http://localhost:5173) in your browser.

## Production build

```bash
npm run build
npm run preview
```

`npm run build` writes a static site to `dist/`. `npm run preview` serves that build locally.

## Project structure

```
tic-tac-toe/
├── index.html          # App shell
├── package.json        # Scripts and dependencies
├── vite.config.js      # Vite configuration
└── src/
    ├── main.jsx        # React entry point
    ├── App.jsx         # Tutorial game (Square, Board, Game)
    └── styles.css      # Tutorial styles
```

## How to play

1. X starts. Click an empty square to place a mark.
2. Players alternate until someone gets three in a row (or the board is full).
3. Use the move list on the right to review earlier turns or start over from the beginning.

## Learn more

- [Tutorial: Tic-Tac-Toe](https://react.dev/learn/tutorial-tic-tac-toe)
- [Vite React guide](https://vite.dev/guide/)
