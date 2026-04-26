import numpy as np
from scipy.stats import truncnorm, beta

def generar_bateria_inicial():
    mu, sigma = 32, 15
    a, b = (0 - mu) / sigma, (100 - mu) / sigma
    dist = truncnorm(a, b, loc=mu, scale=sigma)
    return dist.rvs()

def generar_bateria_final(b_ini):
    alpha, beta_param = 20, 2
    
    while True:
        # Beta.rvs() genera valores en [0, 1], multiplicamos por 100
        b_fin = beta.rvs(alpha, beta_param) * 100
        
        if b_fin >= b_ini:
            return b_fin

for i in range(1, 10):
    bat = generar_bateria_inicial()
    print (bat)
    print(generar_bateria_final(bat))

