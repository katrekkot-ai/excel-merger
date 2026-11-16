"""Локальный рецепт pandas для python-for-android.

Отличия от штатного рецепта p4a (pythonforandroid/recipes/pandas):
  * версия v3.1.1 — первая ветка с официальной поддержкой Python 3.14
    (p4a по умолчанию собирает именно 3.14, а pandas v2.3.0 под неё
    не компилируется: "member reference type 'int' is not a pointer");
  * убран fix_numpy_includes.patch: он написан под meson.build pandas 2.x и
    к 3.x не применяется (hunks FAILED).

Подключается через `p4a.local_recipes = ./p4a-recipes` в buildozer.spec.
"""

from os.path import join

from pythonforandroid.recipe import MesonRecipe


class PandasRecipe(MesonRecipe):
    version = 'v3.1.1'
    url = 'git+https://github.com/pandas-dev/pandas'
    depends = ['numpy', 'libbz2', 'liblzma']
    # meson does not detect venv's cython, so install it explicitly for host python
    hostpython_prerequisites = ["versioneer[toml]", "numpy>=2.0", "Cython<4.0.0a0"]
    patches = []
    python_depends = ['python-dateutil', 'pytz']
    need_stl_shared = True

    def get_recipe_env(self, arch, **kwargs):
        env = super().get_recipe_env(arch, **kwargs)
        # include headers of the numpy build we just compiled for android
        env['NUMPY_INCLUDES'] = join(
            self.ctx.get_python_install_dir(arch.arch), "numpy/_core/include"
        )
        env["PYTHON_INCLUDE_DIR"] = self.ctx.python_recipe.include_root(arch)
        # fixes ImportError: cannot locate symbol "_ZTVSt12length_error"
        # in pandas/_libs/window/*.so
        env['LDFLAGS'] += f' -landroid  -l{self.stl_lib_name}'
        return env

    def build_arch(self, arch):
        super().build_arch(arch)
        self.restore_hostpython_prerequisites(["cython"])


recipe = PandasRecipe()
