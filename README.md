# 股的Morning Project

Welcome to the **股的Morning Project** repository by [Shawn-lung](https://github.com/Shawn-lung/-morning). This project is designed as an investor psychology test tool, combining financial analysis, business analytics, and programming to create interactive tools for assessing individual risk aversion and stock analysis. It integrates Excel VBA with Python to perform data processing, simulation, and optimization tasks.

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Celebrity Investing Strategies](#celebrity-investing-strategies)
- [Collaborators](#collaborators)
- [Code Structure](#code-structure)
- [Python Code Highlights](#python-code-highlights)
- [Note](#note)

## Overview
**Note**: You can view the full presentation [here](https://www.canva.com/design/DAGiHD691Ck/QFBgxmKhyEbXJrc19BDiRw/edit?utm_content=DAGiHD691Ck&utm_campaign=designshare&utm_medium=link2&utm_source=sharebutton).

The 股的Morning Project transforms qualitative financial insights into quantitative investment strategies for investor psychology testing. Rather than focusing solely on asset management, this project evaluates individual risk preferences by simulating investment scenarios and generating personalized investment “scores.”

The project leverages a Python-based genetic algorithm to derive optimal investment formula weights—representing the investment philosophies of renowned celebrity investors—and uses Excel VBA to run Monte Carlo simulations for portfolio performance evaluation and reporting.

This comprehensive analytical tool:
- Downloads, cleans, and processes financial data using Python.
- Normalizes key financial metrics (e.g., ROE, EPS, Gross Margin).
- Translates celebrity investing strategies into mathematical formulas.
- Simulates returns and computes portfolio metrics (average return, variance, standard deviation, Sharpe ratio) via VBA.
- Offers distinct analyses for various risk profiles.
- Generates dynamic reports and visualizations to help users understand their investor profile.

## Features

- **Data Download & Processing**  
  Python scripts automatically download stock data, remove obsolete sheets, and process the data into a standardized format.

- **Genetic Algorithm Optimization**  
  A Python-based genetic algorithm derives optimal weights for investment formulas inspired by celebrity investors. These weights are used to translate qualitative investment strategies into quantitative scores.

- **Simulation & Optimization**  
  Excel VBA runs Monte Carlo simulations to evaluate portfolio performance and optimize weights, calculating essential metrics such as the Sharpe ratio.

- **Risk Profile Analysis**  
  Provides separate analyses for:
  - **風險厭惡 (Risk-Averse / Snow Day)**
  - **風險中立偏厭惡 (Risk-Neutral, Slightly Averse / Cloudy Day)**
  - **風險中立 (Risk-Neutral / Sunny Day)**
  - **風險中立偏愛好 (Risk-Neutral, Slightly Risk-Loving / Thunderstorm Day)**
  - **風險愛好 (Risk-Loving / Lightning Day)**

- **Interactive User Forms & Reporting**  
  Users select a celebrity investor (e.g., Buffet, Graham, O'Shaughnessy, Murphy Score, or a Random option) through an interactive form. Based on their choices, personalized investment scores are generated and visualized using charts (e.g., pie charts showing return percentage breakdown), with the option to export reports as PDFs.

- **Excel VBA & Python Integration**  
  Combines Excel’s VBA for simulations and reporting with Python’s advanced data processing capabilities, including a genetic algorithm for formula optimization and data cleaning.

## Celebrity Investing Strategies

The project translates the investment philosophies of celebrated investors into quantifiable mathematical formulas. These formulas use normalized financial metrics weighted by parameters optimized via a Python genetic algorithm. For example:

- **Buffet Strategy:**  
  Emphasizes ROE, EPS, Gross Margin, and Revenue per Share while penalizing a high P/B Ratio.  
  \[
  \text{Buffet Score} = 0.2808 \times \text{Normalized ROE} + 0.3004 \times \text{Normalized EPS} + 0.16 \times \text{Normalized Gross Margin} + 0.16 \times \text{Normalized Revenue per Share} - 0.0988 \times \text{Normalized P/B Ratio}
  \]

- **Graham Strategy:**  
  Focuses on low PE and P/B ratios along with EPS.  
  \[
  \text{Graham Score} = 0.4286 \times \text{Normalized PE Ratio} + 0.4286 \times \text{Normalized P/B Ratio} + 0.1429 \times \text{Normalized EPS}
  \]

- **O'Shaughnessy Strategy:**  
  Prioritizes EPS and ROE, with a lesser weight on PE Ratio.  
  \[
  \text{O'Shaughnessy Score} = 0.637 \times \text{Normalized EPS} + 0.2583 \times \text{Normalized PE Ratio} + 0.1047 \times \text{Normalized ROE}
  \]

- **Lynch Strategy:**  
  Balances PE Ratio, Revenue per Share, and Gross Margin while reducing the impact of a high P/B Ratio.  
  \[
  \text{Lynch Score} = 0.5825 \times \text{Normalized PE Ratio} + 0.2362 \times \text{Normalized Revenue per Share} + 0.0789 \times \text{Normalized Gross Margin} - 0.1024 \times \text{Normalized P/B Ratio}
  \]

- **Murphy Strategy:**  
  Combines ROE and Operating Margin to deliver an overall performance score.  
  \[
  \text{Murphy Score} = 0.6132 \times \text{Normalized ROE} + 0.5868 \times \text{Normalized Operating Margin}
  \]

These formulas are further optimized by a Python genetic algorithm that fine-tunes the weights based on historical data and consistency criteria, ensuring that the scores accurately reflect the underlying investment philosophy.

## Collaborators

This project is a collaborative effort by members of our dedicated team. Special thanks to the following team members for their contributions:

- Hsiang-I, Lung
- Emily, Guo
- 蕭宇宸
- 蔣伊婷
- 林睿駿
- 陳子鈞
- 蔡宜靜
- 熊彥婷

Their combined expertise in finance, programming, and analytics has been instrumental in developing this investor psychology test tool.

## Code Structure

- **VBA Modules:**  
  - **UserForm Modules:** Manage user interactions (celebrity selection, quiz questions).  
  - **Data Processing Modules:** Handle data import, cleaning, and preparation within Excel.  
  - **Simulation & Optimization Modules:** Execute Monte Carlo simulations for return simulation and portfolio optimization.  
  - **Result Reporting Modules:** Generate charts, copy results, and export PDFs.

- **Python Scripts:**  
  - **Genetic Algorithm & Data Processing:**  
    - The genetic algorithm optimizes the weights for the investment formulas based on historical data and defined celebrity strategies.
    - Python is also used for downloading, cleaning, and normalizing stock data.
  - **Integration:**  
    - Python scripts (e.g., `main.py`) are invoked from VBA, ensuring a seamless flow between data processing and financial simulation.

## Python Code Highlights

- **Genetic Algorithm for Investment Formulas:**  
  The Python script uses a genetic algorithm to optimize the weights for various financial metrics. It:
  - Initializes candidate solutions using a mix of predefined weights and random selections.
  - Uses fitness functions based on the consistency ratio from an AHP model plus a penalty for deviation from initial criteria.
  - Evolves the candidate population to find the optimal weights that best represent a given celebrity investor’s strategy.

- **Stock Scoring Model:**  
  After cleaning and normalizing stock data, the script calculates scores for each stock using the defined strategies, allowing for a quantitative comparison of stocks based on the investment insights derived from celebrity strategies.

## Final Score Calculation

The final score for a stock is derived by applying the optimal weights from the genetic algorithm to its normalized financial metrics. This score serves as the basis for portfolio optimization within the Excel VBA modules, which run Monte Carlo simulations to evaluate risk and performance. Together, these processes enable a robust, data-driven approach to understanding individual investor psychology.

## Note
This repository is solely for admission purposes and is not intended for public use.
