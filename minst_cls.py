import os
import time

import torch
import torch.nn as nn
import torch.utils.data as torchdata
import torchvision
from sklearn.metrics import accuracy_score, f1_score


EPOCH = 3
BATCH_SIZE = 24
LR = 0.001


train_dataset = torchvision.datasets.MNIST(
    root='./mnist',
    download=True,
    train=True,
    transform=torchvision.transforms.ToTensor(),
)
train_loader = torchdata.DataLoader(
    dataset=train_dataset,
    batch_size=BATCH_SIZE,
    shuffle=True
)

val_dataset = torchvision.datasets.MNIST(
    root='./mnist',
    download=False,
    train=False,
    transform=torchvision.transforms.ToTensor(),
)
val_loader = torchdata.DataLoader(
    dataset=val_dataset,
    batch_size=BATCH_SIZE,
    shuffle=False
)
test_x = torch.unsqueeze(val_dataset.data[:500], dim=1).float()
test_y = val_dataset.targets[:500]
print('train_dataset', train_dataset.data.size(), train_dataset.targets.size())
print('test_dataset', val_dataset.data.size(), val_dataset.targets.size())

import matplotlib.pyplot as plt
# plt.imshow(val_dataset.data[0].numpy(), cmap='gray')
# plt.title('%i' % val_dataset.targets[0])
# plt.show()

from matplotlib import cm
from sklearn.manifold import TSNE


class Net(nn.Module):

    def __init__(self):
        super(Net, self).__init__()
        self.level1 = nn.Sequential(
            nn.Conv2d(1, 16, kernel_size=5, stride=1, padding=2),
            nn.BatchNorm2d(16),
            nn.ReLU(),
            nn.MaxPool2d(2)
        )
        self.level2 = nn.Sequential(
            nn.Conv2d(16, 32, kernel_size=3, stride=1, padding=1),
            nn.BatchNorm2d(32),
            nn.ReLU(),
            nn.MaxPool2d(2)
        )
        self.cls = nn.Linear(32 * 7 * 7, 10)

    def forward(self, x):
        x = self.level1(x)
        x = self.level2(x)
        x = x.view(x.size(0), -1)
        output = self.cls(x)
        return output, x


model = Net()
optimizer = torch.optim.Adam(params=model.parameters(), lr=LR)
loss_func = nn.CrossEntropyLoss()


def plot_with_labels(data, labels):
    plt.cla()
    X, Y = data[:, 0], data[:, 1]
    for x, y, s in zip(X, Y, labels):
        c = cm.rainbow(int(255 * s / 9))
        plt.text(x, y, s, backgroundcolor=c, fontsize=9)
    plt.xlim(X.min(), X.max())
    plt.ylim(Y.min(), Y.max())
    plt.title('Visualize last layer')
    plt.show()
    plt.pause(0.01)


plt.ion()
best_f1 = 0
for epoch in range(EPOCH):
    total_train_loss = 0.0
    start_time = time.time()
    model.train()
    for step, (x, y) in enumerate(train_loader):
        output, _ = model(x)
        loss = loss_func(output, y)
        total_train_loss += loss.item()

        optimizer.zero_grad()
        loss.backward()
        optimizer.step()

        if step % 50 == 0 or step == len(train_loader) - 1:
            print('Train Epoch {:04d}, Step {:04d}, Loss {:.4f}, Time {:.4f}'
                  .format(epoch + 1, step + 1, total_train_loss / (step + 1), time.time() - start_time))

            test_output, last_layer = model(test_x)
            test_pred = torch.max(test_output, dim=1)[1].cpu().numpy()
            accuracy = float((test_pred == test_y.numpy()).astype(int).sum()) / float(test_y.size(0))
            print('Train Epoch {:04d}, Step {:04d}, Accuracy {:.4f},'
                  .format(epoch + 1, step + 1, accuracy))
            tsne = TSNE(perplexity=30, n_components=2, init='pca', n_iter=5000)
            plot_only = 500
            low_dim_embs = tsne.fit_transform(last_layer.detach().cpu().numpy()[:plot_only, :])
            labels = test_y.numpy()[:plot_only]
            plot_with_labels(low_dim_embs, labels)

    val_true, val_pred = [], []
    total_val_loss = 0.0
    start_time = time.time()
    model.eval()
    with torch.no_grad():
        for step, (x, y) in enumerate(val_loader):
            output, _ = model(x)
            loss = loss_func(output, y)
            total_val_loss += loss.item()

            if step % 50 == 0 or step == len(train_loader) - 1:
                print('Val Epoch {:04d}, Step {:04d}, Loss {:.4f}, Time {:.4f}'
                      .format(epoch + 1, step + 1, total_val_loss / (step + 1), time.time() - start_time))

            pred = torch.max(output, dim=1)[1]
            val_pred.extend(pred.cpu().numpy().tolist())
            val_true.extend(y.cpu().numpy().tolist())

    val_acc = accuracy_score(val_true, val_pred)
    val_f1 = f1_score(val_true, val_pred, average='micro')
    if best_f1 <= val_f1:
        best_f1 = val_f1
        print('best model, acc {:.4f}, f1 {:.4f}'.format(val_acc, val_f1))
        torch.save(model, 'minst_cls_best.pt')

plt.ioff()

model = torch.load('minst_cls_best.pt')
test_x = torch.unsqueeze(val_dataset.data[:10], dim=1).float()
test_y = val_dataset.targets[:10]
output, _ = model(test_x)
test_pred = torch.max(output, dim=1)[1]
print('test_y', test_y)
print('test_pred', test_pred)
