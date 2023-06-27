from gui import Window

if __name__ == '__main__':
    window =  Window(u"摘要生成工具", 450, 220)
    window.bind_command()
    window.run()
