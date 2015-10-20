# coding: utf-8
import switcher3

window = switcher3.Window("CabinetWClass",None,"explorer")
window.become_child(None,"ShellView")
window.activate_app()
window.debug()
