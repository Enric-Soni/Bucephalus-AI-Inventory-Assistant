BUCEPHALUS AI INVENTORY ASSISTANT

A web-based inventory analysis tool built with Next.js and TypeScript. The application allows users to have inventory workbooks, analyze product and inventory data, and generate actionable inventory insights with an AI chatbot powered by Botpress to explain results.

Requirements
     Node.js
     npm
     Microsoft Excel
Setup

Clone the repository and open the project folder in Terminal.

Install the required dependencies:

     npm install

Start the development server:

     npm run dev

Then open the local address shown in the Terminal.

Excel Add-In

This project also includes an Excel add-in through manifest.xml.

To use the add-in, start the development server first and then sideload manifest.xml into Microsoft Excel.

PROJECT STRUCTURE

app/ — application pages, components, and API routes

lib/ — inventory analysis and processing logic

public/ — images and other public assets

manifest.xml — Excel add-in configuration

READ THIS WRITE-UP FIRST/ — setup instructions, documentation, and example workbooks

NOTES

node_modules is not included in the repository. Run npm install after cloning the project to install all required dependencies.
