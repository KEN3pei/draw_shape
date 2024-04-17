## Getting Started

Welcome to the VS Code Java world. Here is a guideline to help you get started to write Java code in Visual Studio Code.

## Folder Structure

The workspace contains two folders by default, where:

- `src`: the folder to maintain sources
- `lib`: the folder to maintain dependencies

Meanwhile, the compiled output files will be generated in the `bin` folder by default.

> If you want to customize the folder structure, open `.vscode/settings.json` and update the related settings there.

## Dependency Management

The `JAVA PROJECTS` view allows you to manage your dependencies. More details can be found [here](https://github.com/microsoft/vscode-java-dependency#manage-dependencies).

## 使用しているオブジェクト

- [HSSF and XSSF](https://poi.apache.org/components/spreadsheet/)
- そのほかApachPOIがアクセス可能なオブジェクト（[Component Map](https://poi.apache.org/components/index.html)）

## Cache delete

`mvn dependency:purge-local-repository`
`rm -rf ~/.m2/repository/org/apache/maven`

## 実行手順

1. `mvn install`
2. `java -jar target/gs-maven-0.1.0.jar`
