# Tic-Tac-Toe

A complete, runnable Tic-Tac-Toe game based on the [official React tutorial](https://react.dev/learn/tutorial-tic-tac-toe).

This folder is the full app. After you clone the repository, install dependencies here and start the development server.

## Features

- 3×3 board where X and O take turns
- Winner detection with a highlighted winning line
- Draw detection when the board is full and nobody has won
- Move history (“time travel”) so you can jump back to any earlier position
- Reset to start a new game
- Responsive layout: two cards on desktop, stacked cards on small screens

## Prerequisites

| Tool | Version | Why you need it |
| --- | --- | --- |
| [Git](https://git-scm.com/) | any recent version | Clone the repository |
| [Node.js](https://nodejs.org/) | **18 or later** (tested with Node 22) | Runs the JavaScript tooling |
| npm | comes with Node.js | Installs packages and runs scripts |

Check that they are installed:

```bash
git --version
node -v
npm -v
```

`node -v` should print `v18.x.x` or higher. If `node` is missing, install the LTS build from [nodejs.org](https://nodejs.org/) and open a new terminal.

You do **not** need to install React or Vite globally. `npm install` downloads them into this folder.

## Install

### 1. Clone the repository

```bash
git clone https://github.com/anthonyleedu/anthonyleedu.git
cd anthonyleedu
```

The repository is public, so this works while you are signed out of GitHub.

If the `tic-tac-toe` folder is missing, switch to the project branch:

```bash
git checkout cursor/tic-tac-toe-react-tutorial-63df
```

### 2. Enter this project folder

```bash
cd tic-tac-toe
```

If you already cloned the repo and are reading this file locally, skip the clone step and just `cd` here.

### 3. Install dependencies

```bash
npm install
```

This installs React, React DOM, Vite, and the other packages listed in `package.json` into `node_modules/`.

For a lockfile-exact install:

```bash
npm ci
```

## Run the development server

From this `tic-tac-toe` directory:

```bash
npm run dev
```

You should see output similar to:

```text
VITE v8.x.x  ready in xxx ms

➜  Local:   http://localhost:5173/
```

Open [http://localhost:5173](http://localhost:5173) in your browser.

Keep the terminal open while you use the app. Press `Ctrl+C` to stop the server.

If port `5173` is taken, Vite chooses another port and prints that URL. Use the address it prints.

## How to play

1. X starts. Click an empty square to place a mark.
2. Players alternate. The status line shows **Next player**, **Winner**, or **Draw**.
3. Three marks in a row (row, column, or diagonal) wins. The winning squares are highlighted.
4. After a win or a draw, clicks on the board do nothing.
5. Use the **History** list to jump to any previous move, including **Game start**.
6. Click **Reset** to clear the board and start over.

## Production build

```bash
npm run build
```

Vite writes the static site to `dist/`. Preview it:

```bash
npm run preview
```

Open the URL printed in the terminal (usually [http://localhost:4173](http://localhost:4173)).

## Available scripts

| Command | What it does |
| --- | --- |
| `npm install` | Install project dependencies |
| `npm run dev` | Start the Vite development server with hot reload |
| `npm run build` | Build the production files into `dist/` |
| `npm run preview` | Serve the production build locally |
| `npm run lint` | Run Oxlint on the project |

## Project structure

```text
tic-tac-toe/
├── README.md             # Install and run guide
├── package.json          # Scripts and npm dependencies
├── package-lock.json     # Locked dependency versions
├── vite.config.js        # Vite dev server (port 5173) and build config
├── index.html            # HTML shell that loads src/main.jsx
├── public/
│   └── favicon.svg
└── src/
    ├── main.jsx          # Mounts the React app
    ├── App.jsx           # Square, Board, Game, history, and calculateWinner
    └── styles.css        # Layout and visual styles
```

`node_modules/` and `dist/` are created on your machine. They are ignored by Git and are not required in the clone.

## Tech stack

- [React](https://react.dev/) 19 — function components and `useState`
- [Vite](https://vite.dev/) 8 — development server and production bundler
- Tutorial building blocks: `Square`, `Board`, `Game`, and `calculateWinner`

## Troubleshooting

**`Could not read package.json`**  
You are not in this folder. Run `cd tic-tac-toe` from the repository root so `package.json` is in the current directory.

**`tic-tac-toe: No such file or directory`**  
You are on a branch that does not include the app. Run `git checkout cursor/tic-tac-toe-react-tutorial-63df`.

**`node: command not found` or Node is older than 18**  
Install Node.js 18+ from [nodejs.org](https://nodejs.org/), then open a new terminal and run `node -v`.

**`npm install` fails**  
Confirm this directory contains `package.json` and `package-lock.json`. Remove `node_modules` and retry with `npm ci`.

**The browser cannot connect or the page is blank**  
Confirm `npm run dev` is still running and open the Local URL from the terminal output.

**Port already in use**  
Use the fallback localhost URL Vite prints, or stop the other process bound to 5173.

**Edits do not show up**  
Hard-refresh the browser (`Cmd+Shift+R` or `Ctrl+Shift+R`). The dev server reloads when you save files.

## Tutorial credit

https://react.dev/learn/tutorial-tic-tac-toe
