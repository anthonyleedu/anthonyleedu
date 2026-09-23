# Tic-Tac-Toe

This public repository contains a complete, runnable Tic-Tac-Toe web app based on the [official React tutorial](https://react.dev/learn/tutorial-tic-tac-toe).

The full project source lives in the [`tic-tac-toe`](./tic-tac-toe) folder. You can clone this repository, install the dependencies, and start a local development server without any extra configuration.

## Features

- 3×3 board where X and O take turns
- Winner detection with a highlighted winning line
- Draw detection when the board is full and nobody has won
- Move history (“time travel”) so you can jump back to any earlier position
- Reset to start a new game
- Responsive layout: two cards on desktop, stacked cards on small screens

## Prerequisites

Install these before you run the app:

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

`node -v` should print `v18.x.x` or higher. If the `node` command is not found, download the LTS installer from [nodejs.org](https://nodejs.org/) and open a new terminal after installing.

You do **not** need to install React or Vite globally. Those are listed in `tic-tac-toe/package.json` and are installed by `npm install`.

## Install

### 1. Clone the repository

```bash
git clone https://github.com/anthonyleedu/anthonyleedu.git
cd anthonyleedu
```

This repository is public, so cloning works while you are signed out of GitHub.

If the `tic-tac-toe` folder is not on the branch you cloned (for example, you cloned `main` before this project was merged), switch to the project branch:

```bash
git checkout cursor/tic-tac-toe-react-tutorial-63df
```

### 2. Enter the project directory

```bash
cd tic-tac-toe
```

All install and run commands below are meant to be run from this folder.

### 3. Install dependencies

```bash
npm install
```

This reads `package.json` and `package-lock.json`, then downloads React, React DOM, Vite, and the other packages into a local `node_modules` folder. The first install usually takes under a minute.

If you want a clean, lockfile-exact install (the same versions this project was built with):

```bash
npm ci
```

## Run the development server

From `tic-tac-toe/`:

```bash
npm run dev
```

Vite starts a local server. In the terminal you should see something like:

```text
VITE v8.x.x  ready in xxx ms

➜  Local:   http://localhost:5173/
```

Open [http://localhost:5173](http://localhost:5173) in your browser.

Leave the terminal running while you play. Press `Ctrl+C` in that terminal to stop the server.

If port `5173` is already in use, Vite will pick the next free port and print the new URL. Use the URL it prints.

## How to play

1. X starts. Click an empty square to place a mark.
2. Players alternate. The status line shows **Next player**, **Winner**, or **Draw**.
3. Three marks in a row (row, column, or diagonal) wins. The winning squares are highlighted.
4. After a win or a draw, further clicks on the board are ignored.
5. Use the **History** list to jump to any previous move, including **Game start**.
6. Click **Reset** to clear the board and start over.

## Production build

Create an optimized static build:

```bash
cd tic-tac-toe
npm run build
```

The compiled site is written to `tic-tac-toe/dist/`. Preview that build locally:

```bash
npm run preview
```

Then open the URL printed in the terminal (usually [http://localhost:4173](http://localhost:4173)).

## Available scripts

Run these from `tic-tac-toe/`:

| Command | What it does |
| --- | --- |
| `npm install` | Install project dependencies |
| `npm run dev` | Start the Vite development server with hot reload |
| `npm run build` | Build the production files into `dist/` |
| `npm run preview` | Serve the production build locally |
| `npm run lint` | Run Oxlint on the project |

## Project structure

```text
anthonyleedu/
├── README.md                 # This file (install and run guide)
└── tic-tac-toe/              # Complete runnable app
    ├── README.md             # Same setup instructions for the app folder
    ├── package.json          # Scripts and npm dependencies
    ├── package-lock.json     # Locked dependency versions
    ├── vite.config.js        # Vite dev server and build config
    ├── index.html            # HTML shell that loads the React app
    ├── public/
    │   └── favicon.svg
    └── src/
        ├── main.jsx          # React entry point
        ├── App.jsx           # Game, board, history, and winner logic
        └── styles.css        # Layout and visual styles
```

`node_modules/` and `dist/` are generated locally. They are listed in `.gitignore` and are not part of the source you clone.

## Tech stack

- [React](https://react.dev/) 19 (function components and `useState`)
- [Vite](https://vite.dev/) 8 for development and production builds
- The finished tutorial components: `Square`, `Board`, `Game`, and `calculateWinner`

## Troubleshooting

**`tic-tac-toe: No such file or directory`**  
You are not in the repository root, or you are on a branch that does not include the app. Run `ls` and confirm you see `tic-tac-toe`. If you do not, run `git checkout cursor/tic-tac-toe-react-tutorial-63df`.

**`node: command not found` or a Node version below 18**  
Install Node.js 18+ from [nodejs.org](https://nodejs.org/), then open a new terminal and run `node -v` again.

**`npm install` fails**  
Confirm you are inside `tic-tac-toe/` (the folder that contains `package.json`). Delete `node_modules` and try `npm ci`.

**The browser page is blank or cannot connect**  
Make sure `npm run dev` is still running, and open the exact Local URL printed in the terminal.

**Port already in use**  
Use the alternate localhost URL Vite prints, or stop the other process using port 5173.

**Changes do not appear**  
Hard-refresh the browser (`Cmd+Shift+R` or `Ctrl+Shift+R`). The dev server reloads on save.

## License and tutorial credit

Game logic follows the official React Tic-Tac-Toe tutorial:

https://react.dev/learn/tutorial-tic-tac-toe
