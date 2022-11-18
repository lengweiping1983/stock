from sklearn.model_selection import train_test_split
from sklearn import datasets
iris = datasets.load_iris()
x = iris.data
y = iris.target
x_train, x_test, y_train, y_test = train_test_split(x, y)


K = 3
from sklearn.neighbors import KNeighborsClassifier
knn_classifier = KNeighborsClassifier(K, weights="distance")
knn_classifier.fit(x_train, y_train)
y_predict = knn_classifier.predict(x_test)
print('acc:{}'.format(sum(y_predict == y_test) / len(x_test)))
