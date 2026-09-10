Exception in Tkinter callback
Traceback (most recent call last):
  File "C:\Users\lynchd27\AppData\Local\Programs\Python\Python311\Lib\tkinter\__init__.py", line 1948, in __call__
    return self.func(*args)
           ^^^^^^^^^^^^^^^^
  File "C:\Users\lynchd27\Documents\DevPal\new-project\devpal_lite_runner_fixed.py", line 5288, in generate_plan_run_and_validate
    configs = self.collect_configs()
              ^^^^^^^^^^^^^^^^^^^^^^
  File "C:\Users\lynchd27\Documents\DevPal\new-project\devpal_lite_runner_fixed.py", line 5238, in collect_configs
    config = SequenceStepConfig(
             ^^^^^^^^^^^^^^^^^^^
TypeError: SequenceStepConfig() takes no arguments
