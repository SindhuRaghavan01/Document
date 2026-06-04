import numpy as np


def get_policy_gross_loss(condition_losses, location_losses):
    """get total policy gross loss using condition level losses and location
    losses for locations not under any policy conditions

    Parameters
    ----------
    condition_losses : numpy.ndarray
        total gross losses for all conditions under the policy
    location_losses : numpy.ndarray
        total gross losses for all locations under the policy, but not under
        any policy conditions
    """
    return np.sum(condition_losses) + np.sum(location_losses)


def get_policy_deductible(condition_deductibles, location_deductibles):
    """get total policy deductible from condition deductibles and applicable
    location deductibles for locations not under any policy conditions

    Parameters
    ----------
    condition_deductibles : numpy.ndarray
        deductibles for conditions under the policy
    location_deductibles : numpy.ndarray
        applicable deductibles for locations under the policy, but not under
        any policy conditions
    """
    return np.sum(condition_deductibles) + np.sum(location_deductibles)


def get_policy_loss(policy_gross_loss, policy_deductible, policy_applicable_limit,
                    attachment_point, layer_amount, company_share):
    """get final policy losses

    Parameters
    ----------
    policy_gross_loss : numpy.ndarray
        policy gross loss after conditions and location losses
    policy_deductible : numpy.ndarray
        applicable deductible for policy after condition and location deductibles
    policy_applicable_limit : numpy.ndarray
        applicable limit for policy after condition location limits
    attachment_point : numpy.ndarray
        policy attachment point
    layer_amount : numpy.ndarray
        policy layer amount
    company_share : numpy.ndarray
        policy share
    """
