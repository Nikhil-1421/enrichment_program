import setuptools

setuptools.setup(
    name="enrichment_program",
    version="1.0",
    author="Nikhil Datta",
    author_email="nikhildatta@heartsofempowerment.org",
    description="Source Code for Hearts of Empowerment's Enrichment Program Dashboard",
    packages=[
        "apps","src","tests","utils"    
    ],
    install_requires="environment/requirements.txt"
)