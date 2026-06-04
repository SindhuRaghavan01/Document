
def get_groundup_loss(replacement_value: float, damage_ratio: float):
    """Multiply replacement_value with damage ratio to get the ground up loss.
    The DR should be a decimal number in [0,1].

    Parameters
    ----------
    replacement_value : float
        total insured value
    damage_ratio : float
        applicable damage ratio
    """
    return replacement_value * damage_ratio


def get_applicable_deductible(deductible: float, loss: float):
    """calculate effective deductible to apply for site deductibles.
    If the loss is less than the deductible, this loss can still contribute to site/policy deductible.
    So, use the loss as effective deductible in this case.

    Parameters
    ----------
    deductible : float
        coverage deductible
    loss : float
        loss after applying damage ratio
    """
    return min(deductible, loss)


def get_applicable_limit(limit: float, loss: float):
    """calculate effective limit to apply for site limits.

    Parameters
    ----------
    limit : float
        coverage limit
    loss : float
        loss after applying damage ratio
    """
    return min(limit, loss)
