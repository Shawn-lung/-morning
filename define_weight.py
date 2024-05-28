import random
import ahpy
import numpy as np
import pandas as pd
from deap import base, creator, tools, algorithms

# 初始化權重範圍和初始準則
weights_range = [1/9, 1/8, 1/7, 1/6, 1/5, 1/4, 1/3, 1/2, 1, 2, 3, 4, 5, 6, 7, 8, 9]


initial_criteria = {
    ('Normalized ROE', 'Normalized EPS'): 1,
    ('Normalized ROE', 'Normalized Gross Margin'): 2,
    ('Normalized ROE', 'Normalized Revenue per Share'): 2,
    ('Normalized ROE', 'Normalized PB Ratio'): 2,
    ('Normalized EPS', 'Normalized Gross Margin'): 2,
    ('Normalized EPS', 'Normalized Revenue per Share'): 2,
    ('Normalized EPS', 'Normalized PB Ratio'): 3,
    ('Normalized Gross Margin', 'Normalized Revenue per Share'): 1,
    ('Normalized Gross Margin', 'Normalized PB Ratio'): 2,
    ('Normalized Revenue per Share', 'Normalized PB Ratio'): 2
}

initial_weights = list(initial_criteria.values())

# 創建適應度和個體類
creator.create("FitnessMin", base.Fitness, weights=(-1.0,))
creator.create("Individual", list, fitness=creator.FitnessMin)

# 評估函數
def evaluate(individual):
    criteria = {
        ('Normalized ROE', 'Normalized EPS'): individual[0],
        ('Normalized ROE', 'Normalized Gross Margin'): individual[1],
        ('Normalized ROE', 'Normalized Revenue per Share'): individual[2],
        ('Normalized ROE', 'Normalized PB Ratio'): individual[3],
        ('Normalized EPS', 'Normalized Gross Margin'): individual[4],
        ('Normalized EPS', 'Normalized Revenue per Share'): individual[5],
        ('Normalized EPS', 'Normalized PB Ratio'): individual[6],
        ('Normalized Gross Margin', 'Normalized Revenue per Share'): individual[7],
        ('Normalized Gross Margin', 'Normalized PB Ratio'): individual[8],
        ('Normalized Revenue per Share', 'Normalized PB Ratio'): individual[9]
    }
    ahp_model = ahpy.Compare('Criteria', criteria)
    consistency_ratio = ahp_model.consistency_ratio
    
    # 加入懲罰項
    penalty = sum((individual[i] - initial_weights[i]) ** 2 for i in range(len(individual)))
    penalty_weight = 0.1  # 懲罰項的權重，可以調整
    
    return consistency_ratio + penalty_weight * penalty,

# 創建個體
def create_individual():
    return creator.Individual([random.choices([initial_weights[i], random.choice(weights_range)], weights=[0.8, 0.2])[0] for i in range(10)])

# 自定義變異函數
def mutCustom(individual, indpb):
    for i in range(len(individual)):
        if random.random() < indpb:
            individual[i] = random.choices([initial_weights[i], random.choice(weights_range)], weights=[0.8, 0.2])[0]
    return individual,

# 設置DEAP工具箱
toolbox = base.Toolbox()
toolbox.register("individual", create_individual)
toolbox.register("population", tools.initRepeat, list, toolbox.individual)
toolbox.register("mate", tools.cxTwoPoint)
toolbox.register("mutate", mutCustom, indpb=0.2)
toolbox.register("select", tools.selTournament, tournsize=3)
toolbox.register("evaluate", evaluate)

# 運行GA優化
def run_ga_optimization(n_gen=100, pop_size=50, cxpb=0.5, mutpb=0.2):
    random.seed(40)
    pop = toolbox.population(n=pop_size)
    
    # 初始化部分個體為初始準則
    for ind in pop[:int(pop_size * 0.2)]:
        for i, weight in enumerate(initial_weights):
            ind[i] = weight
    
    hof = tools.HallOfFame(1)
    stats = tools.Statistics(lambda ind: ind.fitness.values)
    stats.register("avg", np.mean)
    stats.register("std", np.std)
    stats.register("min", np.min)
    stats.register("max", np.max)
    algorithms.eaSimple(pop, toolbox, cxpb=cxpb, mutpb=mutpb, ngen=n_gen, 
                        stats=stats, halloffame=hof, verbose=True)
    return hof[0]

# 運行GA優化並打印最佳個體
best_individual = run_ga_optimization()

optimal_criteria = {
    ('Normalized ROE', 'Normalized EPS'): best_individual[0],
    ('Normalized ROE', 'Normalized Gross Margin'): best_individual[1],
    ('Normalized ROE', 'Normalized Revenue per Share'): best_individual[2],
    ('Normalized ROE', 'Normalized PB Ratio'): best_individual[3],
    ('Normalized EPS', 'Normalized Gross Margin'): best_individual[4],
    ('Normalized EPS', 'Normalized Revenue per Share'): best_individual[5],
    ('Normalized EPS', 'Normalized PB Ratio'): best_individual[6],
    ('Normalized Gross Margin', 'Normalized Revenue per Share'): best_individual[7],
    ('Normalized Gross Margin', 'Normalized PB Ratio'): best_individual[8],
    ('Normalized Revenue per Share', 'Normalized PB Ratio'): best_individual[9]
}

ahp_model = ahpy.Compare('Criteria', optimal_criteria)
print("Optimal Weights:")
print(ahp_model.report())

# 保存最佳權重到CSV文件
weights_df = pd.DataFrame.from_dict(optimal_criteria, orient='index', columns=['Weight'])
weights_df.to_csv('optimal_weights.csv')

# 應用最佳權重來計算得分（可選）
def calculate_scores(row, weights):
    return sum(row[col] * weight for col, weight in weights.items())

# 示例行和最佳權重
sample_row = {
    'Normalized ROE': 0.5,
    'Normalized EPS': 0.6,
    'Normalized Gross Margin': 0.7,
    'Normalized Revenue per Share': 0.4,
    'Normalized PB Ratio': 0.3
}
optimal_weights = {
    'Normalized ROE': best_individual[0],
    'Normalized EPS': best_individual[1],
    'Normalized Gross Margin': best_individual[2],
    'Normalized Revenue per Share': best_individual[3],
    'Normalized PB Ratio': best_individual[4]
}
score = calculate_scores(sample_row, optimal_weights)
print(f"Sample Score: {score}")
