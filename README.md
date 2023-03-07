# Microsoft Office Add-In for PhonoWriter

## Setting up the environment for the development build

To run the Add-In you'll need to have Node.js installed on your system.

Latest recommended version for Node.js can be downloaded [here](https://nodejs.org/en/).

To check if you already have Node.js installed, type the folloving command in the command line or Terminal:

```
npm -v
```

After the Node.js installation run the following command to install Yeoman Generator:

```
npm install -g yo generator-office
```

Then install Office JavaScript linter by running the following commands:

```
npm install office-addin-lint --save-dev
npm install eslint-plugin-office-addins --save-dev
```

_Please note that you'll also need a code editor (Visual Studio Code, for example) and a Microsoft 365 account_

["Set up your development enviroment" Official MS Documentation](https://learn.microsoft.com/en-gb/office/dev/add-ins/overview/set-up-your-dev-environment?tabs=yeomangenerator)

## Open the Add-In

Clone the repository with Git by running the following command:

```
git clone https://github.com/Tartynn/PhonoWritter.git
```

or alternatively download the project folder by clicking on the `< > Code` -button on the top-right corner and selecting `Download ZIP`.

Open the project on Visual Studio Code, then open the Terminal and run the following command:

```
npm start
```

This will open directly a Word document with the Add-in usually already open on the right side of the screen. If not you can click on “Show Taskpane” at the top right corner to make the Add-in appear.

---

# About the project

The goal of this project is to create a Microsoft Office Add-in for the Application called “PhonoWriter”. PhonoWriter is a writing aid that can help anyone who has difficulty writing sentences or texts. It combines simple word prediction features with complex, innovative and novel algorithms to make writing easier. PhonoWriter is composed of different "word prediction" modules: classic, phonetic, fuzzy, pictographic and sentence.

It is mainly intended for :

- dysorthographics,

- dyslexics,

- dysgraphics,

- allophones.

PhonoWriter is composed of different "word prediction" modules:

**Classic**: usual operation, like on a smartphone, the user enters a few letters, then the module suggests several words beginning with the entered letters. This can save time and accelerate the user's typing speed.

> « bonjo » => « bonjour »

**Phonetic**: suggests words that are phonetically close to the word entered, which allows the user to start the word with a wrong letter or to make big spelling mistakes.

> « ariko » => « haricot »

> « ipotèz » => « hypothèse »

**Pictographic**: displays an image when typing, to allow the distinction of homophones, i.e. to differentiate words that are pronounced the same way but have a different meaning and spelling.

> « mer » => « mère », « maire », « mer »

**Blurring**: to correct inversions or confusions of letters within a word.

> « brobliam » => « problème »

**Sentence**: proposes at the selection of a sentence, corrections to be evaluated in case of doubt on the spelling and/or grammar

> « je vé alla mère avec ma mer » => « je vais à la mer avec ma mère »

**The next word**: anticipates the user's typing, according to their usage, to suggest the most relevant word, before and during typing.

> « Je suis venu » => « je suis venu chez/avec/… »

[PhonoWriter Homepage](https://www.jeanclaudegabus.ch/produits/phonowriter-windows/)
