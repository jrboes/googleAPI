from setuptools import setup, find_packages
import versioneer


with open('requirements.txt', 'r') as f:
    requirements = f.readlines()


setup(
    name='Google API',
    version=versioneer.get_version(),
    cmdclass=versioneer.get_cmdclass(),
    description='Google API tools for drive and sheets',
    author='Jacob Boes',
    author_email='jacobboes@gmail.com',
    packages=find_packages(),
    license='MIT',
    classifiers=[
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python",
        "Programming Language :: Python :: 3",
    ],
    install_requires=requirements,
    include_package_data=True
)
