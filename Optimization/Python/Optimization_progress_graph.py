# animated_line_plot.py
import sys
from random import randint
import os
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
import pandas as pd
from scipy.stats import kurtosis
# create empty lists for the x and y data
x = []
y = []
optimization_id = sys.argv[1]
output_dir = f"output/{optimization_id}/"
progress_file_path = os.path.join(output_dir, "progress.csv")
# create the figure and axes objects
fig, ax = plt.subplots(nrows=2, ncols=3)
print(progress_file_path)
# function that draws each frame of the animation


def animate(i):

    for a in ax:
        for b in a:
            b.clear()
    # a.clear()
    if os.path.exists(progress_file_path):
        df = pd.read_csv(progress_file_path, sep=",")  # .tail(100)

        current_iteration = df["Iteration"].max()
        ax[0][1].plot(df["Iteration"], -df["ExpReturn"] / 1e6)
        ax[0][1].set_title("Expected Return (M)")
        ax[0][1].grid()
        ax[1][1].plot(df["Iteration"], df["TVAR100"] / 1e6)
        ax[1][1].set_title("TVAR 100 (M)")
        ax[1][1].grid()
        # ax[1][0].plot(df["Iteration"], df["PeakCATBondAUM"])
        # ax[1][0].plot(df["Iteration"], df["NonPeakCATBondAUM"])
        ax[1][0].plot(df["Iteration"], df["PortAUM"] / 1e6)
        ax[1][0].set_title("AUM (M)")
        ax[1][0].grid()
        ax[0][0].plot(df["Iteration"], df["FeasibleCount"])
        ax[0][0].set_title("Feasible count")
        ax[0][0].grid()
        ax[0][2].scatter(df["TVAR100"] / 1e6, -df["ExpReturn"] /
                         1e6, c=df["Iteration"] / current_iteration, cmap="Reds")
        ax[0][2].set_title("Exp Return (M) vs TVAR 100 (M)")
        ax[0][2].grid()

        ff_path = os.path.join(output_dir, f"frontier_{current_iteration}.csv")
        if os.path.exists(ff_path):
            ff = pd.read_csv(ff_path)
            k = kurtosis(ff["ExpReturn"])
            ax[1][2].plot(ff["TVAR100"] / 1e6, -ff["ExpReturn"] / 1e6, ".")
            ax[1][2].set_title(f"Frontier - Iteration {current_iteration} - kurt: {k: 0.2f}")
            ax[1][2].grid()
        # ax.plot(x, y)
    # ax.set_xlim([0,20])
    # ax.set_ylim([0,10])


# run the animation
ani = FuncAnimation(fig, animate, frames=20000, interval=500, repeat=False)

plt.show()
