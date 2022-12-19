import pstats


p = pstats.Stats('__init__.prof')
p.sort_stats('cumtime').print_stats()