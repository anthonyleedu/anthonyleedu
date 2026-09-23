# Tic-Tac-Toe

A complete, runnable Tic-Tac-Toe game based on the [official React tutorial](https://react.dev/learn/tutorial-tic-tac-toe).

- 3×3 board with alternating X and O turns
- Winner and draw detection
- Highlighted winning line
- Move history so you can review any previous position

## Install

Install [Node.js](https://nodejs.org/) 18 or later, then:

```bash
git clone https://github.com/anthonyleedu/anthonyleedu.git
cd anthonyleedu/tic-tac-toe
npm install
```

## Run

```bash
npm run dev
```

Open [http://localhost:5173](http://localhost:5173). The game starts in the browser after Vite is ready.

## Production build

```bash
npm run build
npm run preview
```

`npm run build` writes a static site to `dist/`. `npm run preview` serves that build locally.

## Project source

```
tic-tac-toe/
├── index.html
├── package.json
├── package-lock.json
├── vite.config.js
└── src/
    ├── main.jsx      # React entry point
    ├── App.jsx       # Game, board, and history
    └── styles.css
```
