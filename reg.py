import os

import torch
import torch.nn as nn
import matplotlib.pyplot as plt


EPOCH = 300
LR = 0.2

x = torch.unsqueeze(torch.linspace(-1, 1, 100), dim=1)
y = x.pow(2) + 0.2 * torch.rand(x.size())

# plt.scatter(x.numpy(), y.numpy())
# plt.show()


class Net(nn.Module):

    def __init__(self, n_feature, n_hidden, n_output):
        super(Net, self).__init__()
        self.level1 = nn.Linear(n_feature, n_hidden)
        self.level2 = nn.Linear(n_hidden, n_output)

        self.bn = nn.BatchNorm1d(n_hidden)
        self.relu = nn.ReLU()

    def forward(self, x):
        x = self.relu(self.bn(self.level1(x)))
        x = self.level2(x)

        return x


model = Net(1, 10, 1)
optimizer = torch.optim.Adam(params=model.parameters(), lr=LR)
loss_func = nn.MSELoss()

plt.ion()
for epoch in range(EPOCH):
    pred = model(x)
    loss = loss_func(pred, y)

    optimizer.zero_grad()
    loss.backward()
    optimizer.step()

    if epoch % 5 == 0:
        plt.cla()
        plt.scatter(x.numpy(), y.numpy())
        plt.plot(x.numpy(), pred.detach().cpu().numpy(), 'r-', lw=5)
        plt.text(0.5, 0, 'Loss=%.4f' % loss.item(), fontdict={'size': 20, 'color': 'red'})
        plt.pause(0.01)


plt.ioff()
plt.show()







