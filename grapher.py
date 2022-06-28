import matplotlib.pyplot as plt
import numpy as np
from scipy.optimize import curve_fit

def quad(X, a, b, c, d, e, f):
    x,y = X
    return a + b*x + c*y + d*x*x + e*x*y + f*y*y
def fitToQuad (x, y, z):
    p0 = 1, 1, 1, 1, 1, 1
    popt, pcov = curve_fit(quad, (x,y), z, p0)
    return popt, pcov
def sine_function(x, a, b, c):
    return a * np.sin(b * x + c)
def fitToSine (x, y):
    p0 = 1, 1, 1
    popt, pcov = curve_fit(sine_function, x, y, p0)
    return popt, pcov
def getPlot(coefficients, title):
    fig, axes = plt.subplots(nrows=1, ncols=1, figsize=(18, 6))
    x = np.linspace(-4, 4, 100)
    y = np.linspace(-4, 4, 100)
    xv, yv= np.meshgrid(x, y)
    
    radius = 3.9**2
    r = xv*xv+yv*yv
    outside = r > radius
    xv[outside] = float("Nan")
    yv[outside] = float("Nan")
    z = coefficients[0]+coefficients[1]*xv+coefficients[2]*yv+coefficients[3]*xv*xv +coefficients[4]*xv*yv+coefficients[5]*yv*yv
    axes.set_aspect(1)
    plot = axes.pcolor(xv, yv, z, cmap='magma')
    cbar = fig.colorbar(plot, ax=axes)
    title = axes.set_title(title)
    plt.xlabel("X Displacement (Degrees)")
    plt.ylabel("Y Displacement (Degrees)")
    
    plt.show()