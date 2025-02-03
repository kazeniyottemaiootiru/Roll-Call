# 随机点名器
本点名器基于 python 3.12（64-bit） 编写，因此你需要拥有可用的 python 3 的较新版本并且是 64 位的 python 来使用。
## 使用方法
1. 下载源码文件 Excelsetting.py 和 Rollcall.py 文件。
2. 通过 pip 下载 pandas 和 pyinstaller 下载指令见下：
    ```sh
    pip install pandas
    pip install pyinstaller
    ```
3. 将两个源码文件放入同一个文件夹下，然后使用你的 python 的 IDE 打开 Excelsetting.py 并按照提示运行即可。
4. 最后打包完成后在当前路径下的 dist 文件夹下找到生成的 Rollcall.exe 打开后即可运行。

- 注：生成的 Rollcall.exe 在 Window 10 (64-bit) 和 Windows 11 无python环境中均可运行，该代码未对 Linux 做兼容，如要在 Linux 上使用，在代码的 --add-data 参数中的 ; 改为 : 即可。
