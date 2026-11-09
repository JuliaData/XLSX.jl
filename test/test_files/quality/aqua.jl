
#
# Aqua_tests.jl - package-level checks with Aqua.jl: method ambiguities, unbound type
# parameters, undefined exports, stale or uncapped dependencies, type piracy, and
# tasks left running when the package is loaded. All together take about 15 seconds.
#
# `undocumented_names` stays off: the deprecated `chartSchema` and `chartType` aliases
# have no docstrings, deliberately.
#

import Aqua

@testset "Aqua" begin
    Aqua.test_all(XLSX; undocumented_names=false)
end
