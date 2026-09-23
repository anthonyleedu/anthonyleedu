# Tic-Tac-Toe

This public repository includes a complete, runnable Tic-Tac-Toe app based on the [official React tutorial](https://react.dev/learn/tutorial-tic-tac-toe).

The project source is in the [`tic-tac-toe`](./tic-tac-toe) folder.

## Install and run

You need [Node.js](https://nodejs.org/) 18 or later (npm is included).

```bash
git clone https://github.com/anthonyleedu/anthonyleedu.git
cd anthonyleedu/tic-tac-toe
npm install
npm run dev
```

Then open [http://localhost:5173](http://localhost:5173) in your browser.

If you cloned the default `main` branch and the `tic-tac-toe` folder is missing, check out the project branch first:

```bash
git clone https://github.com/anthonyleedu/anthonyleedu.git
cd anthonyleedu
git checkout cursor/tic-tac-toe-react-tutorial-63df
cd tic-tac-toe
npm install
npm run dev
```

## Other commands

```bash
npm run build    # production build
npm run preview  # serve the production build
```

## Project source

```
tic-tac-toe/
├── package.json
├── package-lock.json
├── vite.config.js
├── index.html
├── public/
└── src/
    ├── main.jsx
    ├── App.jsx
    └── styles.css
```
