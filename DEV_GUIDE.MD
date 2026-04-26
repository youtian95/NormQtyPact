# 开发文档

## 开发环境

```bash
pip install -e .
```

## 打包

```bash
pip install build
```

构建 wheel 和 sdist：

```bash
python -m build
```

## 发布

```bash
pip install twine
```

上传到 PyPI：

```bash
python -m twine upload dist/*
```
