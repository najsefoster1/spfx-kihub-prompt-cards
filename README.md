# SPFx KIHub Prompt Cards

This project is a SharePoint Framework (SPFx) solution that provides an interactive prompt library designed to help users quickly generate, copy, and use structured prompts directly within SharePoint.

## Overview

The KIHub Prompt Cards solution was built to make prompt usage more accessible and actionable for users. Instead of presenting prompts as static text, this solution delivers them through an interactive card-based interface that allows users to immediately copy or use prompts in tools like Microsoft Copilot.

The experience is designed to reduce friction, support different skill levels, and encourage adoption of AI-assisted workflows.

## Key Features

- Interactive prompt cards with structured layouts  
- Beginner and Advanced prompt options  
- One-click "Copy Prompt" functionality  
- "Use in Copilot" action for immediate execution  
- Categorization by tool or use case (e.g., SharePoint, Power BI)  
- Clean, modern UI designed for ease of use  

## Business Value

This solution transforms how users interact with prompts by making them immediately usable instead of just informational.

It helps:
- Reduce time spent figuring out how to structure prompts  
- Support both beginner and advanced users  
- Increase adoption of tools like Copilot  
- Standardize prompt quality across teams  
- Improve efficiency and consistency in AI-assisted tasks  

## Tech Stack

- SharePoint Framework (SPFx)  
- TypeScript  
- React  
- Fluent UI  

## Project Structure

src/                SPFx web part or component source code  
config/             Build and configuration files  
sharepoint/         Solution packaging assets  
assets/             Screenshots and visuals  

## Getting Started

npm install  
gulp serve  

## Build and Package

gulp bundle --ship  
gulp package-solution --ship  

## Deployment

1. Upload the .sppkg file to the SharePoint App Catalog  
2. Deploy the solution  
3. Add the web part to your SharePoint page  

## Screenshots

![Prompt Cards](/kihub-prompt-cards.png)

## Author

Najse Foster  
https://github.com/najsefoster1  

## Notes

This project reflects a broader effort to make AI tools more approachable and practical within SharePoint by combining structured guidance with immediate usability.

This project is part of a broader effort to improve how users access AI-related tools, resources, and guidance within SharePoint.
