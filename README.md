# About Markov Chain Generators

A transition matrix can be calculated based on a training sequence (ex. 1, 2, 3). A [Markov Chain Generator](https://github.com/Gagniuc/Markov-Chain-Generator) (MCG) is a prediction machine that uses a transition matrix to generate sequences that are similar to the training sequence. Thus, the output of a MCG mimics the training sequence that led to the values from the transition matrix and the process itself represents a prediction. Moreover, the MCG can also be used to verify the correct operation of the [DPD algorithm](https://github.com/Gagniuc/Discrete-Probability-Detector-in-VB6). Once the DPD algorithm produces a transition matrix (called here the “original” transition matrix) using a training sequence, that transition matrix can be used by the MCG to predict a similar sequence. In turn, the sequence produced by the MCG can be used by the [DPD algorithm](https://github.com/Gagniuc/Discrete-Probability-Detector-JS) to produce a new transition matrix. If the original transition matrix and the transition matrix of the predicted sequence contain close transition probability values, then the DPD algorithm and the MCG machine work as expected. 

# Markov Chains Simulation framework

The application from below is a MCG that uses probability values from a transition matrix to generate text sequences of 10000 letters in length. At each step a new text sequence is analyzed and the letter frequencies are computed. These frequencies are displayed as signals on a graph at each step in order to capture the overall behavior of the MCG. Note that Markov Chains - Simulation framework is made in Visual Basic 6.0 (VB6).

![screenshot](https://github.com/Gagniuc/Markov-Chains-Simulation-framework/blob/main/img/Markov%20Chains%20-%20Simulation%20framework.png?raw=true)

# References

- <i>Paul A. Gagniuc. Markov chains: from theory to implementation and experimentation. Hoboken, NJ,  John Wiley & Sons, USA, 2017, ISBN: 978-1-119-38755-8.</i>
