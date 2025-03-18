# 股的Morning Project

An Interactive Financial Analysis & Recommendation Tool

This repository contains the Python code for the 股的Morning project – a financial analysis platform that integrates real-time stock data, technical indicators, and advanced algorithms to generate personalized investment recommendations. A key focus of this project is the use of Monte Carlo simulation to optimize portfolio weights.

---

## Overview

The 股的Morning Project is designed to transform qualitative financial insights into quantitative investment strategies. The project leverages a Python-based genetic algorithm to derive investment formulas inspired by renowned investors, and applies Monte Carlo simulations to evaluate and optimize portfolio performance.

---

## Key Features

- **Real-Time Data Collection**  
  Automatically fetches key stock metrics such as previous close, today's open, limit up/down prices, current price, and price change.

- **Technical Indicator Integration**  
  Supports multiple indicators (RSI, MACD, Bollinger Bands, and a custom turning-point indicator) using TA-Lib.

- **Quantitative Modeling**  
  Implements a Python-based genetic algorithm to generate optimal investment formulas from historical data and investor-inspired strategies.

- **Monte Carlo Simulation**  
  - Simulates future stock return paths to optimize portfolio weights.
  - Identifies the configuration that maximizes performance metrics (e.g., Sharpe ratio).
  - Provides a statistical basis for evaluating risk and return profiles.

- **User-Friendly Interface (PyQt)**  
  Delivers an interactive graphical interface for data visualization and analysis.

---

## Python Code Highlights

### Genetic Algorithm for Investment Formulas
- **Purpose**  
  Derive optimal investment formulas inspired by renowned investors.
- **Implementation**  
  - Uses a population-based approach where candidate solutions evolve over iterations.
  - Fitness functions are defined by key financial metrics (e.g., ROE, EPS, PB ratio).
  - Best-performing formulas guide portfolio recommendations.

### Monte Carlo Simulation for Portfolio Optimization
- **Purpose**  
  Simulate future returns to optimize portfolio weights and assess risk.
- **Implementation**  
  - **Data Processing:** Cleans and prepares historical stock data using Python (NumPy, pandas).  
  - **Simulation Process:**  
    - Generates 1,000+ simulated return paths for selected stocks.  
    - Evaluates each simulation’s expected return and volatility.  
  - **Optimization:**  
    - Identifies the portfolio with the highest Sharpe ratio.  
    - Outputs optimal weights, visualized via charts (e.g., pie charts).

---

## Project Structure

